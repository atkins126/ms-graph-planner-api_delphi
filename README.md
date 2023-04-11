# ms-graph-planner-api_delphi

This is a small module and cli to access the Microsoft Planner API from Delphi.

## Usage

### Module

The module uses the [MicrosoftApiAuthentication](https://github.com/MeroFuruya/MicrosoftApiAuthentication)-Package to authenticate against the Microsoft Graph API. You can use the `MsGraphPlannerApi`-Module like this:

```delphi
uses
  MicrosoftPlanner,
  MicrosoftApiAuthenticator;

var
  Authenticator: TMsAuthenticator;
  Planner: TMsPlanner;
begin
  Authenticator := TMsAuthenticator.Create(...) // see MicrosoftApiAuthentication-repo
  Planner := TMsPlanner.Create(Authenticator);
end;
```

### CLI

```bash
Usage: planner_cli [command] [option]

Commands:
  list
  create
  update
  delete
  help

Items:
  -g, --Group [<Group id>]
  -p, --Planner [<Board id>]
  -b, --Bucket [<Bucket id>]
  -t, --Task [<Task id>]

Options:
  -h, --Help
  -q, --Quiet
  --Fields "<field>=<value>,<field>=<value>"
  --Debug

Authentication:
  Can be provided as command line parameters or as environment variables.
  Command line parameters have precedence.

  Command line parameters:
    --TenantID "<your tenant id>"
    --ClientID "<your client id>"
    --RedirectURI "<your redirect uri>"
    --RedirectPort "<your redirect port>"
    --Scope "<scope>,<scope>,<scope>"

  Environement Variables:   (should be self explanatory)
    PLANNER_CLI_TENANT_ID
    PLANNER_CLI_CLIENT_ID
    PLANNER_CLI_REDIRECT_URI
    PLANNER_CLI_REDIRECT_PORT
    PLANNER_CLI_SCOPE
```
