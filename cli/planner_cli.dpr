program planner_cli;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  System.StrUtils,
  System.DateUtils,
  System.Generics.Collections,
  MicrosoftApiAuthenticator,
  helpers;

procedure printHelp();
begin
  WriteLn('' + sLineBreak
    + 'Usage: planner_cli [command] [options]' + sLineBreak
    + '' + sLineBreak
    + 'Commands:' + sLineBreak
    + '  list' + sLineBreak
    + '  help' + sLineBreak
    + '' + sLineBreak
    + 'Options:' + sLineBreak
    + '  -t --TenantID="<your tenant id>"' + sLineBreak
    + '  -c, --ClientID="<your client id>"' + sLineBreak
    + '  -r, --RedirectURI="<your redirect uri>"' + sLineBreak
    + '  -p, --RedirectPort="<your redirect port>"' + sLineBreak
    + '  -s, --Scope="<scope>,<scope>,<scope>"' + sLineBreak
    + '  -q, --Verbose' + sLineBreak
    + '' + sLineBreak
    + 'Environement Variables:' + sLineBreak
    + '  PLANNER_CLI_TENANT_ID' + sLineBreak
    + '  PLANNER_CLI_CLIENT_ID' + sLineBreak
    + '  PLANNER_CLI_REDIRECT_URI' + sLineBreak
    + '  PLANNER_CLI_REDIRECT_PORT' + sLineBreak
    + '  PLANNER_CLI_SCOPE' + sLineBreak
    + '' + sLineBreak
  );
end;

function getOptions(): TDictionary<string, string>;
var
  i: Integer;
  s: string;
begin
  Result := TDictionary<string, string>.Create();
  if ParamCount > 1 then
  for i := 2 to ParamCount do
  begin
    s := ParamStr(i);
    if s.StartsWith('-') then
    begin
      // remove leading dash/es
      s := s.Remove(0, 1);
      if s.StartsWith('-') then
        s := s.Remove(0, 1);
      // parse option
      if s.Contains('=') then
      begin
        Result.Add(s.Substring(0, s.IndexOf('=')), s.Substring(s.IndexOf('=') + 1));
      end
      else
      begin
        Result.Add(s, '');
      end;
    end;
  end;
end;

var
  Verbose: Boolean;
  Command: String;
  Options: TDictionary<string, string>;

  TENANT_ID: string;
  CLINET_ID: string;
  REDIRECT_URI: string;
  REDIRECT_PORT: Integer;
  SCOPE: TArray<string>;

  HELPERS: THelpers;


procedure list();
var
  Groups: TList<THelperGroup>;
  Group: THelperGroup;
  Planner: THelperPlanner;
  OutMsg: TArray<string>;

  procedure msg(s: string);
  begin
    OutMsg := OutMsg + [s];
  end;
begin
  // get groups
  Groups := TList<THelperGroup>.Create(HELPERS.getAllPlanners());
  // generate output
  for Group in Groups do
  begin
    msg('- Group: ' + Group.DisplayName);
    msg('  ID: ' + Group.ID);
    msg('  Description: ' + Group.Description);
    msg('  Created: ' + Group.CreatedDateTime);

    for Planner in Group.Planners do
    begin
      msg('    - Planner: ' + Planner.Title);
      msg('      ID: ' + Planner.ID);
      msg('      Owner: ' + Planner.Owner);
      msg('      Created: ' + Planner.CreatedDateTime);
    end;
  end;
  // print output
  if Verbose then
  begin
    writeln(string.Join(sLineBreak, OutMsg));
  end;
  Groups.Free();
end;

begin
  if ParamCount = 0 then
  begin
    printHelp();
    Exit;
  end;

  // get command
  Command := ParamStr(1).ToLower();

  if Command = 'help' then
  begin
    printHelp();
    Exit;
  end;

  // get options
  Options := getOptions();

  // check for verbose
  Verbose := not (Options.ContainsKey('q') or Options.ContainsKey('QUIET'));

  // get tenant id
  if Options.ContainsKey('t') then
    TENANT_ID := Options['t']
  else if Options.ContainsKey('TenantID') then
    TENANT_ID := Options['TenantID']
  else if GetEnvironmentVariable('PLANNER_CLI_TENANT_ID') <> '' then
    TENANT_ID := GetEnvironmentVariable('PLANNER_CLI_TENANT_ID')
  else
  begin
    writeln('Tenant ID not set. Use -t or --TenantID option or set PLANNER_CLI_TENANT_ID environment variable.');
    Exit;
  end;

  // get client id
  if Options.ContainsKey('c') then
    CLINET_ID := Options['c']
  else if Options.ContainsKey('ClientID') then
    CLINET_ID := Options['ClientID']
  else if GetEnvironmentVariable('PLANNER_CLI_CLIENT_ID') <> '' then
    CLINET_ID := GetEnvironmentVariable('PLANNER_CLI_CLIENT_ID')
  else
  begin
    writeln('Client ID not set. Use -c or --ClientID option or set PLANNER_CLI_CLIENT_ID environment variable.');
    Exit;
  end;

  // get redirect uri
  if Options.ContainsKey('r') then
    REDIRECT_URI := Options['r']
  else if Options.ContainsKey('RedirectURI') then
    REDIRECT_URI := Options['RedirectURI']
  else if GetEnvironmentVariable('PLANNER_CLI_REDIRECT_URI') <> '' then
    REDIRECT_URI := GetEnvironmentVariable('PLANNER_CLI_REDIRECT_URI')
  else
  begin
    writeln('Redirect URI not set. Use -r or --RedirectURI option or set PLANNER_CLI_REDIRECT_URI environment variable.');
    Exit;
  end;

  // get redirect port
  if Options.ContainsKey('p') then
    REDIRECT_PORT := StrToInt(Options['p'])
  else if Options.ContainsKey('RedirectPort') then
    REDIRECT_PORT := StrToInt(Options['RedirectPort'])
  else if GetEnvironmentVariable('PLANNER_CLI_REDIRECT_PORT') <> '' then
    REDIRECT_PORT := StrToInt(GetEnvironmentVariable('PLANNER_CLI_REDIRECT_PORT'))
  else
  begin
    writeln('Redirect Port not set. Use -p or --RedirectPort option or set PLANNER_CLI_REDIRECT_PORT environment variable.');
    Exit;
  end;

  // get scope
  if Options.ContainsKey('s') then
    SCOPE := Options['s'].Split([','])
  else if Options.ContainsKey('Scope') then
    SCOPE := Options['Scope'].Split([','])
  else if GetEnvironmentVariable('PLANNER_CLI_SCOPE') <> '' then
    SCOPE := GetEnvironmentVariable('PLANNER_CLI_SCOPE').Split([','])
  else
  begin
    writeln('Scope not set. Use -s or --Scope option or set PLANNER_CLI_SCOPE environment variable.');
    Exit;
  end;

  HELPERS := THelpers.New(TENANT_ID, CLINET_ID, REDIRECT_URI, REDIRECT_PORT, SCOPE, Verbose);

  // PARSE COMMANDS
  if Command = 'list' then
  begin
    list();
  end
  else
  begin
    writeln('Unknown command: ' + Command);
    printHelp();
  end;

  HELPERS.Free;
  options.Free;
end.
