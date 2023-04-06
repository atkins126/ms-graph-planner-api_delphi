﻿program planner_cli;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  System.StrUtils,
  System.DateUtils,
  System.Generics.Collections,
  MicrosoftApiAuthenticator,
  MicrosoftPlanner,
  helpers,
  listing;

procedure printHelp();
begin
  WriteLn('' + sLineBreak
    + 'Usage: planner_cli [command] [option]' + sLineBreak
    + '' + sLineBreak
    + 'Commands:' + sLineBreak
    + '  planner' + sLineBreak
    + '  list' + sLineBreak
    + '  help' + sLineBreak
    + '' + sLineBreak
    + 'Options:' + sLineBreak
    + '  -i, --GetInfo <- get info about the specified item' + sLineBreak
    + '  -c, --Create  <- create a new item' + sLineBreak
    + '  -u, --Update  <- update an existing item' + sLineBreak
    + '  -d, --Delete  <- delete an existing item' + sLineBreak
    + sLineBreak
    + 'Items:' + sLineBreak
    + '  -g, --Group [<Group id>] <- used to specify a group' + sLineBreak
    + '  -p, --Planner [<Board id>] <- used to specify a planner' + sLineBreak
    + '  -b, --Bucket [<Bucket id>] <- used to specify a bucket' + sLineBreak
    + '  -t, --Task [<Task id>] <- used to specify a task' + sLineBreak
    + sLineBreak
    + 'Other Options:' + sLineBreak
    + '  --TenantID "<your tenant id>"' + sLineBreak
    + '  --ClientID "<your client id>"' + sLineBreak
    + '  --RedirectURI "<your redirect uri>"' + sLineBreak
    + '  --RedirectPort "<your redirect port>"' + sLineBreak
    + '  --Scope "<scope>,<scope>,<scope>"' + sLineBreak
    + '  -q, --Quiet' + sLineBreak
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

function resolveAlias(param: string): string;
begin
  if IndexText(param, ['i', 'GetInfo']) <> -1 then
    Result := 'GetInfo'
  else if IndexText(param, ['c', 'Create']) <> -1 then
    Result := 'Create'
  else if IndexText(param, ['u', 'Update']) <> -1 then
    Result := 'Update'
  else if IndexText(param, ['d', 'Delete']) <> -1 then
    Result := 'Delete'
  else if IndexText(param, ['g', 'Group']) <> -1 then
    Result := 'Group'
  else if IndexText(param, ['p', 'Planner']) <> -1 then
    Result := 'Planner'
  else if IndexText(param, ['b', 'Bucket']) <> -1 then
    Result := 'Bucket'
  else if IndexText(param, ['t', 'Task']) <> -1 then
    Result := 'Task'
  else if IndexText(param, ['q', 'Quiet']) <> -1 then
    Result := 'Quiet'
  else
    Result := param;
end;

begin
  Result := TDictionary<string, string>.Create();
  if ParamCount > 1 then
  begin
    //iterate over all params
    i := 0;
    while i < ParamCount do
    begin
      s := ParamStr(i);
      if s.StartsWith('-') then
      begin
        // remove first leading dash
        s := s.Remove(0, 1);
        // check if param is short or long
        if s.StartsWith('-') then
        begin
          // long
          s := s.Remove(0, 1);
          // check if param has value
          if s.Contains('=') then
          begin
            Result.Add(resolveAlias(s.Substring(0, s.IndexOf('='))), s.Substring(s.IndexOf('=') + 1));
          end
          // check if next param is value
          else if (i + 1) <= ParamCount then
          begin
            if not ParamStr(i + 1).StartsWith('-') then
            begin
              Result.Add(resolveAlias(s), ParamStr(i + 1));
              Inc(i);
            end
            else
            begin
              Result.Add(resolveAlias(s), '');
            end;
          end
          else
          begin
            Result.Add(resolveAlias(s), '');
          end;
        end
        else
        begin
          // short
          // check if param has value
          if s.Contains('=') then
          begin
            Result.Add(resolveAlias(s.Substring(0, s.IndexOf('='))), s.Substring(s.IndexOf('=') + 1));
          end
          // check if next param is value
          else if (i + 1) <= ParamCount then
          begin
            if not ParamStr(i + 1).StartsWith('-') then
            begin
              Result.Add(resolveAlias(s), ParamStr(i + 1));
              Inc(i);
            end
            else
            begin
              Result.Add(resolveAlias(s), '');
            end;
          end
          else
          begin
            Result.Add(resolveAlias(s), '');
          end;
        end;
      end;
      Inc(i);
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

  listing: TListing;

  //Groups: TList<THelperGroup>;

// function getId(param: string): string;
// var
//   group: TMsPlannerGroup;
//   planner: TMsPlannerPlanner;
//   ParamValue: string;
// begin
//   Result := '';
//   if param = '' then
//     Exit;
//   
//   if Options.TryGetValue(param, ParamValue) then
//   begin
//     // check if param is id
//     for group in Groups do
//     begin
//       for planner in group.Planners do
//       begin
//         if ((planner.ID = param) or (planner.Title = param)) and (IndexText(param, ['p', 'Planner']) <> -1) then
//         begin
//           Result := planner.ID;
//           Exit;
//         end;
//       end;
//     end;
//   end;
// end;

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
  Verbose := not (Options.ContainsKey('q') or Options.ContainsKey('Quiet'));

  // get tenant id
  if Options.ContainsKey('TenantID') then
    TENANT_ID := Options['TenantID']
  else if GetEnvironmentVariable('PLANNER_CLI_TENANT_ID') <> '' then
    TENANT_ID := GetEnvironmentVariable('PLANNER_CLI_TENANT_ID')
  else
  begin
    writeln('Tenant ID not set. Use --TenantID option or set PLANNER_CLI_TENANT_ID environment variable.');
    Exit;
  end;

  // get client id
  if Options.ContainsKey('ClientID') then
    CLINET_ID := Options['ClientID']
  else if GetEnvironmentVariable('PLANNER_CLI_CLIENT_ID') <> '' then
    CLINET_ID := GetEnvironmentVariable('PLANNER_CLI_CLIENT_ID')
  else
  begin
    writeln('Client ID not set. Use --ClientID option or set PLANNER_CLI_CLIENT_ID environment variable.');
    Exit;
  end;

  // get redirect uri
  if Options.ContainsKey('RedirectURI') then
    REDIRECT_URI := Options['RedirectURI']
  else if GetEnvironmentVariable('PLANNER_CLI_REDIRECT_URI') <> '' then
    REDIRECT_URI := GetEnvironmentVariable('PLANNER_CLI_REDIRECT_URI')
  else
  begin
    writeln('Redirect URI not set. Use --RedirectURI option or set PLANNER_CLI_REDIRECT_URI environment variable.');
    Exit;
  end;

  // get redirect port
  if Options.ContainsKey('RedirectPort') then
    REDIRECT_PORT := StrToInt(Options['RedirectPort'])
  else if GetEnvironmentVariable('PLANNER_CLI_REDIRECT_PORT') <> '' then
    REDIRECT_PORT := StrToInt(GetEnvironmentVariable('PLANNER_CLI_REDIRECT_PORT'))
  else
  begin
    writeln('Redirect Port not set. Use --RedirectPort option or set PLANNER_CLI_REDIRECT_PORT environment variable.');
    Exit;
  end;

  // get scope
  if Options.ContainsKey('Scope') then
    SCOPE := Options['Scope'].Split([','])
  else if GetEnvironmentVariable('PLANNER_CLI_SCOPE') <> '' then
    SCOPE := GetEnvironmentVariable('PLANNER_CLI_SCOPE').Split([','])
  else
  begin
    writeln('Scope not set. Use --Scope option or set PLANNER_CLI_SCOPE environment variable.');
    Exit;
  end;

  HELPERS := THelpers.New(TENANT_ID, CLINET_ID, REDIRECT_URI, REDIRECT_PORT, SCOPE, Verbose);

  // get groups
  // Groups := TList<THelperGroup>.Create(HELPERS.getAllPlanners());s

  // PARSE COMMANDS
  if Command = 'list' then
  begin
    listing := Tlisting.Create(Options, HELPERS.Planner);
    WriteLn(listing.Text);
    listing.Free;
  end
  else
  begin
    writeln('Unknown command: ' + Command);
    printHelp();
  end;

  HELPERS.Free;
  options.Free;
end.