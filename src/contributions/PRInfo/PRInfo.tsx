import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CoreRestClient, WebApiTeam } from "azure-devops-extension-api/Core";
import { GitServiceIds, IVersionControlRepositoryService } from "azure-devops-extension-api/Git/GitServices";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { Page } from "azure-devops-ui/Page";
import { GitRestClient, GitPullRequest, PullRequestStatus, GitPullRequestSearchCriteria } from "azure-devops-extension-api/Git";
import { CommonServiceIds, getClient, IProjectPageService } from "azure-devops-extension-api";
import { showRootComponent } from "../../Common";
import { GitRepository, IdentityRefWithVote } from "azure-devops-extension-api/Git/Git";
import { ITableItem } from "./TableData";
import { getPieChartInfo, getStackedBarChartInfo, stackedChartOptions, BarChartSize, getDurationBarChartInfo, getPullRequestsCompletedChartInfo, ITeamBarChartData, ITeamChartData } from "./ChartingInfo";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Card } from "azure-devops-ui/Card";
import { ObservableArray, ObservableValue } from "azure-devops-ui/Core/Observable";
import { Toast } from "azure-devops-ui/Toast";
import { ZeroData } from "azure-devops-ui/ZeroData";
import { Spinner, SpinnerSize } from "azure-devops-ui/Spinner";
import * as statKeepers from "./statKeepers";
import * as chartjs from "chart.js";
import { Doughnut, Bar } from 'react-chartjs-2';
import { Dropdown } from "azure-devops-ui/Dropdown";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { Observer } from "azure-devops-ui/Observer";
import { DropdownSelection } from "azure-devops-ui/Utilities/DropdownSelection";
import { TeamMember } from "azure-devops-extension-api/WebApi/WebApi";

interface IRepositoryServiceHubContentState {
    repository: GitRepository | null;
    exception: string;
    isToastVisible: boolean;
    isToastFadingOut: boolean;
    foundCompletedPRs: boolean;
    doneLoading: boolean;
    teamsChecked: Map<string, boolean>;
    allTeamsChecked: boolean;
}

class RepositoryServiceHubContent extends React.Component<{}, IRepositoryServiceHubContentState> {
    private itemProvider: ObservableArray<ITableItem | ObservableValue<ITableItem | undefined>>;
    private toastRef: React.RefObject<Toast> = React.createRef<Toast>();
    private totalDuration: number = 0;
    private durationDisplayObject: statKeepers.IPRDuration;
    private targetBranches: statKeepers.INameCount[] = [];
    private branchDictionary: Map<string, statKeepers.INameCount>;

    private teamsDictionary: Map<string, statKeepers.ITeamWithMembers>;

    private approverList: ObservableValue<statKeepers.IReviewWithVote[]>;
    private approverDictionary: Map<string, statKeepers.IReviewWithVote>;

    private approvalTeamDictionary: Map<string, statKeepers.IReviewWithVote>;
    private totalReviewsByTeam: statKeepers.IReviewWithVote[];
    private reviewsByTeam: Map<string, statKeepers.IReviewWithVote[]>;
    public readonly noReviewerText: string = "No Reviewer";

    public myBarChartDims: BarChartSize;
    public PRCount: number = 0;
    private readonly TOP1000_Selection_ID = "36500";
    private readonly dayMilliseconds: number = (24 * 60 * 60 * 1000);
    private completedDate: ObservableValue<Date>;
    private displayText: ObservableValue<string>;
    private prList: GitPullRequest[] = [];
    private rawPRCount: number = 0;
    private dateSelection: DropdownSelection;
    private durationSlices: statKeepers.IDurationSlice[] = [];
    private dateSelectionChoices = [
        { text: "Last 7 Days", id: "7" },
        { text: "Last 14 Days", id: "14" },
        { text: "Last 30 Days", id: "30" },
        { text: "Last 60 Days", id: "60" },
        { text: "Last 90 Days", id: "90" },
        { text: "Top 1000 PRs", id: this.TOP1000_Selection_ID }
    ];

    constructor(props: {}) {
        super(props);
        this.itemProvider = new ObservableArray<ITableItem | ObservableValue<ITableItem | undefined>>(this.getTableItemProvider([]).value);
        this.state = { repository: null, exception: "", isToastFadingOut: false, isToastVisible: false, foundCompletedPRs: true, doneLoading: false, teamsChecked: new Map(), allTeamsChecked: true };
        this.durationDisplayObject = { days: 0, hours: 0, minutes: 0, seconds: 0, milliseconds: 0 };

        this.myBarChartDims = { height: 250, width: 500 };

        this.dateSelection = new DropdownSelection();
        this.dateSelection.select(1);
        this.completedDate = new ObservableValue<Date>(this.getDateForSelectionIndex(1));
        this.displayText = new ObservableValue<string>("Completed Since " + this.completedDate.value.toLocaleDateString());

        this.branchDictionary = new Map<string, statKeepers.INameCount>();
        this.approvalTeamDictionary = new Map<string, statKeepers.IReviewWithVote>();
        this.totalReviewsByTeam = [];
        this.reviewsByTeam = new Map<string, statKeepers.IReviewWithVote[]>();
        this.approverDictionary = new Map<string, statKeepers.IReviewWithVote>();
        this.approverList = new ObservableValue<statKeepers.IReviewWithVote[]>([]);
        this.teamsDictionary = new Map<string, statKeepers.ITeamWithMembers>();
        this.initCollectionValues()
    }

    private initCollectionValues() {
        this.totalDuration = 0;
        this.PRCount = 0;
        this.approvalTeamDictionary.clear();
        this.totalReviewsByTeam = [];
        this.reviewsByTeam.clear();
        this.approverDictionary.clear();
        this.branchDictionary.clear();
        this.approverDictionary.set(this.noReviewerText, { name: this.noReviewerText, value: 0 });
        this.approverList.value = [];
        this.targetBranches = [];
    }

    private getDateForSelectionIndex(ndx: number): Date {
        let dateOffset: number = 0;
        if (this.dateSelectionChoices.length >= ndx) {
            dateOffset = Number.parseInt(this.dateSelectionChoices[ndx].id);
        }
        let RetDate: Date = new Date(new Date().getTime() - (dateOffset * this.dayMilliseconds));

        return RetDate;
    }

    public async componentDidMount() {
        await SDK.init();
        try {
            const repoSvc = await SDK.getService<IVersionControlRepositoryService>(GitServiceIds.VersionControlRepositoryService);
            var repository = await repoSvc.getCurrentGitRepository();
            var exception = "";

            if (repository) {
                this.setState({ repository: repository });
                await this.LoadData();

                if (this.rawPRCount < 1) {
                    this.setState({ foundCompletedPRs: false, isToastFadingOut: false, isToastVisible: false, exception: "", doneLoading: true });
                }
                else {
                    this.setState({ foundCompletedPRs: true, isToastFadingOut: false, isToastVisible: false, exception: "", doneLoading: true });
                }
            }
        }
        catch (ex) {
            if (ex instanceof Error) {
                exception = " Error Retrieving Pull Requests -- " + ex.toString();
                this.toastError(exception);
            }
        }
    }

    private onSelect = (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>) => {
        this.completedDate.value = new Date((new Date().getTime() - (Number.parseInt(item.id) * this.dayMilliseconds)))
        if (item.id == this.TOP1000_Selection_ID) {
            this.displayText.value = "Top 1000";
        }
        else {
            this.displayText.value = "Completed Since " + this.completedDate.value.toLocaleDateString();
        }
        this.approverList.value = [];
        this.handleDateChange();
    };

    private GetTableDataFunctions(prList: GitPullRequest[]): ArrayItemProvider<ITableItem> {
        if (prList) {
            let prTableList = this.getPullRequestRows(prList);
            let prTableArrayObj = this.getTableItemProvider(prTableList);
            return prTableArrayObj;
        }
        else {
            this.setState({ isToastVisible: true, exception: "The List of Pull Requests was not provided when attempting to build the table objects" });
            return new ArrayItemProvider([]);
        }
    }

    private async LoadData() {
        if (this.state.repository) {
            if (this.teamsDictionary.size === 0) {
                await this.retrieveAllMembers(this.state.repository);
            }
            if (this.prList.length === 0) {
                let prList = await this.retrievePullRequestRowsFromADO(this.state.repository.id);
                this.prList = prList.sort(statKeepers.ComparePRClosedDate);
                this.rawPRCount = this.prList.length;
            }
            this.GetTableDataFunctions(this.prList);
            this.AssembleData();
            this.durationSlices = statKeepers.getPRDurationSlices(this.prList);
        }
        else {
            this.setState({ isToastVisible: true, exception: "The Repository ID was not found when attempting to load data!" });
        }
    }

    /// Handle check/uncheck the box for specific team
    private handleCheckedTeamsChange = (e: { target: { name: string; checked: boolean; }; }) => {
        const item = e.target.name;
        const isChecked = e.target.checked;
        this.setState(prevState => {
            let isAllTeamsChecked = true;
            const newTeamsChecked = prevState.teamsChecked.set(item, isChecked);
            newTeamsChecked.forEach(value => {
                isAllTeamsChecked = isAllTeamsChecked && value;
            });
            return {
                teamsChecked: newTeamsChecked,
                allTeamsChecked: isAllTeamsChecked
            }
        });
    };

    /// Handle the check/uncheck the box of all teams
    private handleCheckedAllTeamsChange = (e: { target: { checked: boolean; }; }) => {
        const newTeamsChecked: Map<string, boolean> = new Map();
        const isChecked = e.target.checked;
        this.state.teamsChecked.forEach((_, teamName) => newTeamsChecked.set(teamName, isChecked));
        this.setState(_ => ({
            teamsChecked: newTeamsChecked,
            allTeamsChecked: isChecked
        }));
    }

    private async handleDateChange() {
        this.setState({ doneLoading: false });
        if (this.state.repository) {
            this.LoadData();
        }
        this.setState({ doneLoading: true });
    }

    private AssembleData() {
        try {
            let tempapproverList: statKeepers.IReviewWithVote[] = [];
            let averageOpenTime = 0;
            if (this.PRCount > 0) {
                averageOpenTime = this.totalDuration / this.PRCount;
            }

            this.branchDictionary.forEach((thisBranchItem) => {
                this.targetBranches.push(thisBranchItem);
            });
            this.approverDictionary.forEach((value) => {
                //we will only put the "No Reviewer" item in if we had a PR with no reviewer
                if (value.name == this.noReviewerText) {
                    if (value.value > 0) {
                        tempapproverList.push(value);
                    }

                }
                else {
                    tempapproverList.push(value);
                }

            })

            //sort the lists
            this.approverList.value = tempapproverList.sort(statKeepers.CompareReviewWithVoteByValue);
            this.targetBranches = this.targetBranches.sort(statKeepers.CompareINameCountByValue);

            this.durationDisplayObject = statKeepers.getMillisecondsToTime(averageOpenTime);
        }
        catch (ex) {
            this.toastError("Assembling data: " + ex);
        }
        this.setState({ doneLoading: true });
    }

    public AddRowItem(item: ITableItem) {
        const asyncRow = new ObservableValue<ITableItem | undefined>(undefined);
        this.itemProvider.push(asyncRow);

        asyncRow.value =
        {
            createdBy: item.createdBy,
            prCreatedDate: item.prCreatedDate,
            prCompleteDate: item.prCompleteDate,
            sourceBranch: item.sourceBranch,
            targetBranch: item.targetBranch,
            id: item.id,
            prOpenTime: item.prOpenTime,
            status: item.status,
            reviewerCount: item.reviewerCount
        };
    }

    /// Retrieve all repository teams with associated members
    public async retrieveAllMembers(repository: GitRepository) {
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();
        let projectName = project?.name;
        if (!projectName) {
            // If service doesn't provide the project information, use host page and repository url to guess it
            const hostName = SDK.getHost().name;
            const urlSplit = repository.url.split('/');
            projectName = urlSplit.find(split => !!split && split !== hostName && split !== "_git") || urlSplit[urlSplit.length - 1];
        }
        let teams = await this.retrieveTeams(projectName);
        let teamsChecked: Map<string, boolean> = new Map();
        await teams.forEach(async (team) => {
            let members = (await this.retrieveTeamMembers(projectName!, team)).map((member) => member.identity.displayName);
            this.teamsDictionary.set(team.name, { name: team.name, members });
            teamsChecked.set(team.name, true);
        });
        this.setState({ teamsChecked })
    }

    /// Retrieve the teams related to the project
    public async retrieveTeams(projectId: string): Promise<WebApiTeam[]> {
        const client = getClient(CoreRestClient);
        return client.getTeams(projectId);
    }

    /// Retrieve the members of the team for a given project
    public async retrieveTeamMembers(projectId: string, team: WebApiTeam): Promise<TeamMember[]> {
        const client = getClient(CoreRestClient);
        return client.getTeamMembersWithExtendedProperties(projectId, team.id);
    }

    ///
    public async retrievePullRequestRowsFromADO(repositoryId: string): Promise<GitPullRequest[]> {
        let searchCriteria: GitPullRequestSearchCriteria = { status: PullRequestStatus.Completed, includeLinks: false, creatorId: "", reviewerId: "", repositoryId: "", sourceRefName: "", targetRefName: "", sourceRepositoryId: "" };
        const client = getClient(GitRestClient);
        let prList = client.getPullRequests(repositoryId, searchCriteria, undefined, undefined, undefined, 1000);
        return prList;
    }

    ///
    public getPullRequestRows(prList: GitPullRequest[]): ITableItem[] {

        let rows: ITableItem[] = [];
        try {
            if (prList) {
                this.initCollectionValues();
                prList.forEach((value) => {
                    if (value.closedDate >= this.completedDate.value) {
                        let PROpenDuration = value.closedDate.valueOf() - value.creationDate.valueOf();

                        let thisPR: ITableItem = { createdBy: value.createdBy.displayName, prCreatedDate: value.creationDate, prCompleteDate: value.closedDate, sourceBranch: value.sourceRefName, targetBranch: value.targetRefName, id: value.pullRequestId.toString(), prOpenTime: PROpenDuration, status: value.status.toString(), reviewerCount: value.reviewers.length };

                        this.AddPRDurationToTotalDuration(value);
                        this.AddPRTargetBranchToStat(value);

                        this.AddPRReviewerToStat(value);

                        rows.push(thisPR);
                        this.AddRowItem(thisPR);

                        if (!this.state.foundCompletedPRs) {
                            this.setState({ foundCompletedPRs: true, repository: this.state.repository, isToastFadingOut: false, isToastVisible: false, exception: "", doneLoading: true });

                        }
                        if (!this.state.doneLoading) {
                            this.setState({ doneLoading: true });
                        }
                        this.PRCount += 1;
                    }
                });
                // compute reviews by team
                this.ConsolidateReviewsByTeam();
            }
            else {
                this.setState({ doneLoading: true });
            }
        }
        catch (ex) {
            if (ex instanceof Error) {
                let exception = " Error Retrieving Pull Requests -- " + ex.toString();
                this.toastError("Getting Rows: " + exception);
            }
        }
        return rows;
    }

    /// Consolidate the reviews by team to get reviews by team and reviews inside each teams
    private ConsolidateReviewsByTeam() {
        this.teamsDictionary.forEach((teamWithMembers) => {
            let teamName = teamWithMembers.name;

            let teamReviewsScores: statKeepers.IReviewWithVote[] = [];
            let totalTeamReviewScore: statKeepers.IReviewWithVote = { name: teamName, value: 0 };
            teamWithMembers.members.forEach((member) => {
                let memberScore: statKeepers.IReviewWithVote = this.approverDictionary.get(member) || { name: member, value: 0 };
                totalTeamReviewScore.value += memberScore.value;
                teamReviewsScores.push(memberScore);
            });
            teamReviewsScores = teamReviewsScores.sort(statKeepers.CompareReviewWithVoteByValue)
            this.totalReviewsByTeam.push(totalTeamReviewScore);
            this.reviewsByTeam.set(teamName, teamReviewsScores);
        });
        this.totalReviewsByTeam = this.totalReviewsByTeam.sort(statKeepers.CompareReviewWithVoteByValue)
    }

    private AddPRDurationToTotalDuration(thisPR: GitPullRequest) {
        //get the milliseconds that this PR Was open
        let thisPRDuration = thisPR.closedDate.valueOf() - thisPR.creationDate.valueOf();
        this.totalDuration += thisPRDuration;
    }

    private AddPRTargetBranchToStat(thisPR: GitPullRequest) {
        let branchnameOnly = thisPR.targetRefName.replace("refs/heads/", "")
        let branch = branchnameOnly;
        if (branchnameOnly.split('release/').length > 1) {
            branch = branchnameOnly.split('release/')[1];
        }

        if (this.branchDictionary.has(branch)) {
            let thisref = this.branchDictionary.get(branch);
            if (thisref) {
                thisref.value = thisref.value + 1;
            }

        }
        else {
            this.branchDictionary.set(branch, { name: branch, value: 1 });
        }
    }

    private AddPRReviewerToStat(thisPR: GitPullRequest) {
        if (thisPR.reviewers.length > 0) {
            thisPR.reviewers.forEach(value => {
                if (!value.isContainer) {
                    // individual approver
                    this.AddPRIdentityToStat(value);
                }
            });
        }
        else {
            let thisref = this.approverDictionary.get(this.noReviewerText);
            if (thisref) {
                thisref.value = thisref.value + 1;
            }
        }
    }

    private AddPRIdentityToStat(thisValue: IdentityRefWithVote) {
        let thisID = thisValue.displayName;
        let thisName = thisValue.displayName;
        if (this.approverDictionary.has(thisID)) {
            let thisApprover = this.approverDictionary.get(thisID);
            if (thisApprover) {
                thisApprover.value = thisApprover.value + 1;
                this.approverDictionary.set(thisID, thisApprover);
            }
        }
        else {

            let newVoteStat: statKeepers.IReviewWithVote = { name: thisName, value: 1 };
            this.approverDictionary.set(thisID, newVoteStat);
        }
    }

    public getTableItemProvider(prRows: ITableItem[]): ArrayItemProvider<ITableItem> {
        return new ArrayItemProvider<ITableItem>(
            prRows.map((item: ITableItem) => {
                const newItem = Object.assign({}, item);
                return newItem;
            })
        );
    }

    private toastError(toastText: string) {
        this.setState({ isToastVisible: true, isToastFadingOut: false, exception: toastText, repository: this.state.repository })
    }

    public render(): JSX.Element {
        let isToastVisible = this.state.isToastVisible;
        let foundCompletedPRs = this.state.foundCompletedPRs;
        let doneLoading = this.state.doneLoading;

        let teamNames: string[] = []
        this.state.teamsChecked.forEach((_, teamName) => {
            teamNames.push(teamName);
        });

        let targetBranchChartData = getPieChartInfo(this.targetBranches);
        let reviewerPieChartData = getPieChartInfo(this.approverList.value);
        let reviewerBarChartData = getStackedBarChartInfo(this.approverList.value, this.noReviewerText);
        let smallNumberOfReviewers = reviewerPieChartData.labels.length < 7

        // Global teams charts computation
        let filteredReviewsByTeam = this.totalReviewsByTeam.filter(teamTotalReviews => !!this.state.teamsChecked.get(teamTotalReviews.name))
        let allTeamsBarChartData = getStackedBarChartInfo(filteredReviewsByTeam);
        let allTeamsPieChartData = getPieChartInfo(filteredReviewsByTeam);

        // Team charts computation
        let teamsBarChartData: ITeamBarChartData[] = [];
        let teamPieChartData: ITeamChartData[] = []
        filteredReviewsByTeam.forEach(totalReview => {
            const team = totalReview.name;
            const teamReviews = this.reviewsByTeam.get(totalReview.name);
            if (teamReviews) {
                teamsBarChartData.push({ team, chart: getStackedBarChartInfo(teamReviews) });
                teamPieChartData.push({ team, chart: getPieChartInfo(teamReviews) });
            }
        });
        let maximumVotes = Math.max(...teamsBarChartData.map(teamChart => Math.max(...teamChart.chart.datasets.map(dataSet => Math.max(...dataSet.data, 0)), 0)), 0);

        let teamsBarChartOptions: chartjs.ChartOptions = {
            scales: {
                yAxes: [{ stacked: true, ticks: { beginAtZero: true, max: maximumVotes } }],
                xAxes: [{ stacked: true }]
            },
        }
        let durationTrenChartData = getDurationBarChartInfo(this.durationSlices);
        let closedPRChartData = getPullRequestsCompletedChartInfo(this.durationSlices);
        if (doneLoading) {
            if (!foundCompletedPRs) {
                return (
                    <Page className="sample-hub flex-grow">
                        <Header title="Repository PR Stats" titleSize={TitleSize.Large} />

                        <ZeroData
                            primaryText="No Completed Pull Requests found in this Repository"
                            secondaryText={
                                <span>
                                    This report is designed to give you stats and information about the Pull Request completions in your repository, it will begin providing data as you begin completing Pull Requests in this repository
                                </span>
                            }
                            imageAltText="Bars"
                            imagePath={"./emptyPRList.png"}
                        />
                    </Page>
                );
            }
            else {
                return (
                    <Page className="flex-grow prinfo-hub">
                        <Header title="Repository PR Stats" titleSize={TitleSize.Large} />
                        <div>
                            <div className="flex-row">
                                <div className="flex-column">
                                    <span className="flex-cell" style={{ minWidth: "max-content" }}>
                                        Show Pull Requests Completed within: <span style={{ minWidth: "5px" }} />
                                        <Dropdown
                                            ariaLabel="Basic"
                                            placeholder="Select an Option"
                                            width={500}
                                            items={this.dateSelectionChoices}
                                            selection={this.dateSelection}
                                            onSelect={this.onSelect}
                                        />
                                    </span>
                                </div>
                            </div>
                            {/* Select teams to display in the statistics */}
                            {teamNames.length > 1 &&
                                <div className="flex-row">
                                    <div className="flex-column">
                                        <span className="flex-cell" style={{ minWidth: "max-content" }}>
                                            Select teams for statistics: <span style={{ minWidth: "5px" }} />
                                            <input
                                                type="checkbox"
                                                checked={!!this.state.allTeamsChecked}
                                                onChange={this.handleCheckedAllTeamsChange}
                                            />
                                            All teams
                                            {teamNames.map(teamName => (
                                                <div>
                                                    <input
                                                        type="checkbox"
                                                        name={teamName}
                                                        checked={!!this.state.teamsChecked.get(teamName)}
                                                        onChange={this.handleCheckedTeamsChange}
                                                    />
                                                    {teamName}
                                                    <span style={{ minWidth: "5px" }} />
                                                </div>
                                            ))}
                                        </span>
                                    </div>
                                </div>
                            }
                            <div className="flex-row">
                                <div className="flex-column" style={{ minWidth: "350px" }}>
                                    <div className="flex-row">
                                        <Card titleProps={{ text: this.displayText.value }} >
                                            <div className="flex-cell" style={{ flexWrap: "wrap", minWidth: "350px" }}>
                                                <div className="flex-column" style={{ minWidth: "310px" }} key={1}>
                                                    <div className="body-m secondary-text flex-center" style={{ minWidth: "350px", textAlign: "center" }}>Count</div>
                                                    <div className="title-m flex-center" style={{ minWidth: "350px", textAlign: "center" }}>{this.PRCount}</div>
                                                </div>
                                            </div>
                                        </Card>
                                    </div>
                                    <div className="flex-row">
                                        <div className="flex-cell flex-grow" style={{ minWidth: "350px" }}>
                                            <Card titleProps={{ text: "Closed Pull Requests" }}>
                                                <div className="flex-cell" style={{ minWidth: "315px" }}>
                                                    <table>
                                                        <tbody>
                                                            <tr><td>
                                                                <div style={{ minWidth: "315px" }}><Bar data={closedPRChartData} height={200}></Bar></div>
                                                            </td></tr>
                                                            <tr><td>
                                                                <div className="body-xs" style={{ minWidth: "315px" }}>Trends for the last year (max last 1000 PRs)</div>
                                                            </td></tr>
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </Card>
                                        </div>
                                    </div>
                                </div>
                                <div className="flex-column" style={{ minWidth: "350px" }}>
                                    <div className="flex-row">
                                        <Card titleProps={{ text: "Average Time Pull Requests are Open" }}>
                                            <div className="flex-cell" style={{ flexWrap: "wrap", textAlign: "center", minWidth: "350px" }}>
                                                <div className="flex-column" style={{ minWidth: "70px" }} key={1}>
                                                    <div className="body-m secondary-text">Days</div>
                                                    <div className="title-m primary-text flex-center">{this.durationDisplayObject.days.toString()}</div>
                                                </div>
                                                <div className="flex-column" style={{ minWidth: "70px" }} key={2}>
                                                    <div className="body-m secondary-text">Hours</div>
                                                    <div className="title-m primary-text flex-center">{this.durationDisplayObject.hours.toString()}</div>
                                                </div>
                                                <div className="flex-column" style={{ minWidth: "70px" }} key={3}>
                                                    <div className="body-m secondary-text">Minutes</div>
                                                    <div className="title-m primary-text flex-center">{this.durationDisplayObject.minutes.toString()}</div>
                                                </div>
                                                <div className="flex-column" style={{ minWidth: "70px" }} key={4}>
                                                    <div className="body-m secondary-text">Seconds</div>
                                                    <div className="title-m primary-text flex-center">{this.durationDisplayObject.seconds.toString()}</div>
                                                </div>
                                            </div>
                                        </Card>
                                    </div>
                                    <div className="flex-row">
                                        <div className="flex-cell" style={{ minWidth: "350px" }}>
                                            <Card titleProps={{ text: "Open Time Trends (2 week interval)" }}>
                                                <div className="flex-cell" style={{ minWidth: "315px" }}>
                                                    <table>
                                                        <tbody>
                                                            <tr><td>
                                                                <div className="flex-cell" style={{ minWidth: "315px" }}><Bar data={durationTrenChartData} height={200}></Bar></div>
                                                            </td></tr>
                                                            <tr><td>
                                                                <div className="flex-cell body-xs" style={{ minWidth: "315px" }}>Trends for the last year (max last 1000 PRs)</div>
                                                            </td></tr>
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </Card>
                                        </div>
                                    </div>
                                </div>
                                <div className="flex-column" style={{ minWidth: "350px" }}>
                                    <Card className="flex-grow" titleProps={{ text: "Target Branches" }}>
                                        <div className="flex-row" style={{ flexWrap: "wrap" }}>
                                            <table>
                                                <thead>
                                                    <tr>
                                                        <td></td>
                                                        <td style={{ alignContent: "center", textAlign: "center", minWidth: "85px" }}>Count</td>
                                                        <td style={{ alignContent: "center", textAlign: "center", minWidth: "85px" }}>Percent</td>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {this.targetBranches.map((items, index) => (
                                                        <tr>
                                                            <td className="body-m secondary-text">{items.name}</td>
                                                            <td className="body-m primary-text flex-center" style={{ alignContent: "center", textAlign: "center", minWidth: "85px" }}>{items.value}</td>
                                                            <td className="body-m primary-text flex-center" style={{ alignContent: "center", textAlign: "center", minWidth: "85px" }}>{(items.value / this.PRCount * 100).toFixed(2)}%</td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                    </Card>
                                </div>
                                <div className="flex-column" style={{ minWidth: "450" }}>
                                    <Card className="flex-grow">
                                        <div className="flex-row" style={{ minWidth: "450px" }}>
                                            <Doughnut data={targetBranchChartData} height={200}>
                                            </Doughnut>
                                        </div>
                                    </Card>
                                </div>
                            </div>
                            <div className="flex-row">
                                <div className="flex-column" style={{ minWidth: "350px" }}>
                                    <Card className="flex-grow" titleProps={{ text: "PR Code Reviewers" }}>
                                        <div className="flex-row" style={{ flexWrap: "wrap" }}>
                                            <table>
                                                <thead>
                                                    <tr>
                                                        <td></td>
                                                        <td style={{ alignContent: "center", textAlign: "center", minWidth: "60px" }}>Count</td>
                                                        <td style={{ alignContent: "center", textAlign: "center", minWidth: "100px" }}>Percent of PRs</td>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <Observer selectedItem={this.approverList}>
                                                        {(props: { selectedItem: statKeepers.IReviewWithVote[] }) => {
                                                            return (
                                                                <>
                                                                    {props.selectedItem.map((items, index) => (
                                                                        <tr>
                                                                            <td className="body-m secondary-text flex-center">{items.name}</td>
                                                                            <td className="body-m primary-text flex-center" style={{ alignContent: "center", textAlign: "center", minWidth: "60px" }}>{items.value}</td>
                                                                            <td style={{ alignContent: "center", textAlign: "center", minWidth: "100px" }}>{(items.value / this.PRCount * 100).toFixed(2)}%</td>
                                                                        </tr>
                                                                    ))}
                                                                </>
                                                            )
                                                        }
                                                        }
                                                    </Observer>
                                                </tbody>
                                            </table>
                                        </div>
                                    </Card>
                                </div>
                                {/* Pie chart when there are not too many reviewers */}
                                {smallNumberOfReviewers &&
                                    <div className="flex-column" style={{ minWidth: "500px" }}>
                                        <Card className="flex-grow">
                                            <div className="flex-row flex-grow flex-cell" style={{ minWidth: "500px", height: "220" }}>
                                                <Doughnut data={reviewerPieChartData} height={220}></Doughnut>
                                            </div>
                                        </Card>
                                    </div>
                                }
                                <div className="flex-column">
                                    <Card>
                                        <div className="flex-row" style={{ minWidth: smallNumberOfReviewers ? 400 : 800, height: smallNumberOfReviewers ? 400 : 800 }}>
                                            <Bar data={reviewerBarChartData} options={stackedChartOptions} height={300}></Bar>
                                        </div>
                                    </Card>
                                </div>
                            </div>

                            {/* Bar charts by team */}
                            {allTeamsBarChartData.labels.length &&
                                <div className="flex-row">
                                    {teamsBarChartData.map(barChartData => (
                                        <div className="flex-column" style={{ minWidth: "450px" }}>
                                            <Card className="flex-grow" titleProps={{ text: barChartData.team }}>
                                                <div className="flex-row" style={{ minWidth: 400, height: "300" }}>
                                                    <Bar data={barChartData.chart} options={teamsBarChartOptions} height={300}></Bar>
                                                </div>
                                            </Card>
                                        </div>
                                    ))}
                                </div>
                            }

                            {/* Pie charts by team when there are not too many teams */}
                            {allTeamsBarChartData.labels.length && teamPieChartData.length < 4 &&
                                <div className="flex-row">
                                    {teamPieChartData.map(pieChartData => (
                                        <div className="flex-column" style={{ minWidth: "300px" }}>
                                            <Card className="flex-grow" titleProps={{ text: pieChartData.team }}>
                                                <div className="flex-row flex-grow flex-cell" style={{ minWidth: "300px", height: "160" }}>
                                                    <Doughnut data={pieChartData.chart} height={220}></Doughnut>
                                                </div>
                                            </Card>
                                        </div>
                                    ))}
                                </div>
                            }

                            {/* Charts for global teams */}
                            {allTeamsBarChartData.labels.length > 1 &&
                                <div className="flex-row">
                                    <div className="flex-column" style={{ minWidth: "450px" }}>
                                        <Card className="flex-grow" titleProps={{ text: "Total approvals per team" }}>
                                            <div className="flex-row" style={{ minWidth: 400, height: "300" }}>
                                                <Bar data={allTeamsBarChartData} options={stackedChartOptions} height={300}></Bar>
                                            </div>
                                        </Card>
                                    </div>

                                    <div className="flex-column" style={{ minWidth: "500px" }}>
                                        <Card className="flex-grow">
                                            <div className="flex-row flex-grow flex-cell" style={{ minWidth: "500px", height: "220" }}>
                                                <Doughnut data={allTeamsPieChartData} height={220}></Doughnut>
                                            </div>
                                        </Card>
                                    </div>
                                </div>
                            }
                        </div>

                        {isToastVisible && (
                            <Toast
                                ref={this.toastRef}
                                message={this.state.exception}
                                callToAction="OK"
                                onCallToActionClick={() => { this.setState({ isToastFadingOut: true, isToastVisible: false, exception: "", repository: this.state.repository }) }}
                            />
                        )}
                    </Page>
                );
            }
        }
        else { //else not done loading, so show spinner
            return (
                <Page className="flex-grow">
                    <Header title="Repository PR Stats" titleSize={TitleSize.Large} />
                    <Card className="flex-grow flex-center bolt-table-card" contentProps={{ contentPadding: true }}>
                        <div className="flex-cell">
                            <Spinner label="Loading ..." size={SpinnerSize.large} />
                        </div>
                    </Card>
                </Page>
            );
        }
    }
}

showRootComponent(<RepositoryServiceHubContent />);