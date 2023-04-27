# ms-graph-planner-api_delphi

This is a small module and cli to access the Microsoft Planner API from Delphi.

Recommended Scope for the cli and the module:

```Text
GroupMember.Read.All, Tasks.ReadWrite
```

## 1 CLI

### 1.1 Installation

Just download the latest release from the [release page](https://github.com/MeroFuruya/ms-graph-planner-api_delphi/releases/latest).
Execute `planner_cli.exe` in a command prompt. If you want to use the CLI from anywhere, add the directory to your `PATH` environment variable.

**TIP:** if you dont want to type the info needed for authentication everytime, you can just make a powershell script that sets the environment variables and then calls the CLI. For example:

```powershell
$env:PLANNER_CLI_TENANT_ID = "<your tenant id>"
$env:PLANNER_CLI_CLIENT_ID = "<your client id>"
$env:PLANNER_CLI_REDIRECT_URI = "<your redirect uri>"
$env:PLANNER_CLI_REDIRECT_PORT = "<your redirect port>"
$env:PLANNER_CLI_SCOPE = "<scope>,<scope>,<scope>"
planner_cli.exe $args
```

Or

```powershell
planner_cli.exe $args --TenantID "<your tenant id>" --ClientID "<your client id>" --RedirectURI "<your redirect uri>" --RedirectPort "<your redirect port>" --Scope "<scope>,<scope>,<scope>"
```

### 1.2 Usage

For help, just type `planner_cli.exe --help` or `planner_cli.exe <command> --help`.

#### 1.2.1 listing

```powershell
planner_cli.exe list
```

This will list all the plans in your tenant you are allowed to see.

The structure is always the following:

```Text
  ┌ Groups (-g, --Group)
  │
  ├─┬ Plans (-p, --Plan)
  │ │
  │ ├─┬ Buckets (-b, --Bucket)
  │ │ │
  ╵ ╵ └── Tasks (-t, --Task)
```

#### 1.2.2 Listing without an ID

| Option | Description |
| ------ | ----------- |
| `-g`, `--Group` | List all groups you are allowed to see. |
| `-p`, `--Plan` | List all groups and their plans. |
| `-b`, `--Bucket` | List all groups, their plans and the buckets in their plans. |
| `-t`, `--Task` | List all groups, their plans, the buckets in their plans and the tasks in their buckets. |

#### 1.2.3 Listing with an ID

| Option | Description |
| ------ | ----------- |
| `-g`, `--Group` | List the specified Group. |
| `-p`, `--Plan` | List the specified Plan. |
| `-b`, `--Bucket` | List the specified Bucket. |
| `-t`, `--Task` | List the specified Task. |

Now if you want to list all buckets in a specific plan, you can do it like this:

```powershell
planner_cli.exe list --Planner <plan id> --Bucket
```

You could also list all tasks in a specific bucket:

```powershell
planner_cli.exe list --Bucket <bucket id> --Task
```

Or all tasks in a specific plan:

```powershell
planner_cli.exe list --Planner <plan id> --Task
```

This doesnt work "the other way around" unfortunately. (what i mean ist `--Planner --Task <task id>` wont work)

## 2 Module

### 2.1 Installation (with [boss](https://github.com/hashload/boss))

```powershell
boss install MeroFuruya/ms-graph-planner-api_delphi
```

*Notice: boss currently has a "bug" that prevents it from compiling the cli. You can still use the module though. [#139](https://github.com/HashLoad/boss/pull/139) fixes this.*

### 2.2 Installation (without boss)

Just download the latest source code release from the [release page](https://github.com/MeroFuruya/ms-graph-planner-api_delphi/releases/latest) and add all contents of the the `src` directory to your project.
Then go to [MeroFuruya/MicrosoftApiAuthentication](https://github.com/MeroFuruya/MicrosoftApiAuthentication#installation) and follow the installation instructions there.

### 2.3 Usage

First you need to set up your authentication as described [here](https://github.com/MeroFuruya/MicrosoftApiAuthentication#usage).

Then you can set create a new instance of `TMsPlanner` providing the `TMsGraphAuthenticator` you created before.

```delphi
uses
  MicrosoftApiAuthenticator,
  MicrosoftPlanner;

var
  Autenticator: TMsAuthenticator;
  Planner: TMsPlanner;
begin
  // Create the authenticator
  .
  .
  .
  // Create the planner-api-wrapper
  Planner := TMsPlanner.Create(Authenticator);

  // Do stuff with the planner

  Planner.Free;
  Authenticator.Free;
end;
```

#### 2.3.1 Basic concepts

This framework uses records to represent and store data from the planner api. This has the benefit of not needing to worry about freeing the objects what makes it easier to work with the data.

The downside is that you cant inherit one record from another one what makes it impossible to have dedicated objects for different information stages (e.g. `TMsPlannerGroup` and `TMsPlannerGroupWithPlans`).

The records are also not very "smart" and dont have any methods. They are just containers for the data.

That means that you have to do some things yourself. For example if you want to get all plans in a group, you have to do it like this:

```delphi
var
  Task: TMsPlannerTask;
begin

  // set the id of the task you want to get
  Task.ID := '<task id>';
  
  // Get the task
  Planner.GetTask(Task);
  
  // Now the task contains information you can use
  Writeln('Title:', Task.Title);
end;
```

#### 2.3.2 Listing

```delphi
var
  Plan: TMsPlannerPlanner;
  i: integer;
begin

  // set the id of the plan you want to get
  Plan.ID := '<plan id>';

  // Get all buckets in the plan
  Planner.GetBuckets(Plan);
  
  for i := 0 to length(Plan.Buckets) - 1 do
  begin
    // Now the bucket contains information you can use
    Writeln('Name:', Plan.Buckets[i].Name);
  end;
end;
```

#### 2.3.3 Creating

At the time, only `tasks` and `buckets` can be created.

```delphi
var
  Bucket: TMsPlannerBucket;
begin
  // set the id of the plan you want to create the bucket in
  Bucket.PlanID := '<plan id>';

  // Create a new bucket
  Bucket.Name := 'New Bucket';
  Planner.CreateBucket(Bucket);
  
  // The bucket now is filled with all the information from the api, just like when you "get" it
  Writeln('Name:', Bucket.Name);
end;
```

To get a list of fields you can fill, look at `help` of the cli.

```powershell
planner_cli.exe create --help
```

#### 2.3.4 Updating

At the time, only `tasks`, `buckets` and `planner-categorys` (`PlannerDetails`) can be updated.

Only fill the fields you want to update.

```delphi
var
  Bucket: TMsPlannerBucket;
begin
  // set the id of the bucket you want to update
  Bucket.ID := '<bucket id>';

  // Update the bucket
  Bucket.Name := 'New Bucket Name';
  Planner.UpdateBucket(Bucket);
end;
```

To get a list of fields you can fill, look at `help` of the cli.

```powershell
planner_cli.exe update --help
```

To update categorys of a task, add those you want to change to the `Categorys` array of the task and then update the task.
If you want to remove a category, set its `enabled` property to `false`.
The id of the category must match the id of the category in the plan.
The Name of the category can only be changed by changing the name of the category in the plan.

```delphi
var
  Task: TMsPlannerTask;
begin
  // set the id of the task you want to update
  Task.ID := '<task id>';

  // Add a category
  SetLength(Task.Categorys, 1);
  Task.Categorys[0].ID := '<category id>';
  Task.Categorys[0].Enabled := true;

  // Update the task
  Planner.UpdateTask(Task);
end;
```

To update the checklist of a task, add those you want to change to the `Checklist` array of the task and then update the task.
If you want to remove a checklist item, set its `IsDeleted` property to `true`.

```delphi
var
  Task: TMsPlannerTask;
begin
  // set the id of the task you want to update
  Task.ID := '<task id>';

  // Add a checklist item
  SetLength(Task.Checklist, 1);
  Task.Checklist[0].Title := 'New Checklist Item';
  Task.Checklist[0].IsDeleted := false;

  // Update the task
  Planner.UpdateTask(Task);
end;
```

#### 2.3.5 Deleting

At the time, only `tasks` and `buckets` can be deleted.

```delphi
var
  Bucket: TMsPlannerBucket;
begin
  // set the id of the bucket you want to delete
  Bucket.ID := '<bucket id>';

  // Delete the bucket
  Planner.DeleteBucket(Bucket);
end;
```

### 3 Further information

You can always go take a look at the [Microsoft Graph Planner API documentation](https://docs.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0) to get more information about the api. That is a pretty good resource.
