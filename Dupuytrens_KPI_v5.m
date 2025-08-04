% -------------------------------------------------------------------------
% This script is designed to query Dupuytrens data for Dr. Eaton, whose
% request is "Is there an automated way for me to see how many new 
% enrollees there have been during each phase? To develop my enrollment 
% and follow-up outreach efforts, I want to track the number of new 
% enrollees and phase participants in real time to track how well my 
% own outreach efforts are.". THIS SCRIPT ADDS THE CALCULATIONS OF THE 
% PERCENT OF PARTICIPANTS THAT COMPLETE EACH SURVEY PAGE. 02/03/2025
% -------------------------------------------------------------------------

% --- Save data to a spreadsheet where the rows correspond to the query and 
% the columns correspond to the date the queries were run ---

% Get a list of files in the current directory to check for 'Dupuytrens KPI.xlsx'
currentDir = dir(pwd);
if ~any(contains({currentDir.name}, 'Dupuytrens KPI.xlsx')) % if existing spreadsheet doesn't exist, create one from scratch
    outputData = cell(length(queryList) + 1, 3);
    outputData{1,1} = 'Title';
    outputData{1,2} = 'Description';
    outputData{1,3} = string(datetime('now','Format','yyyy/MM/dd, HH:mm')); % insert current datetime at top of column 3
    for queryNum = 1:length(queryList)
        outputData{queryNum+1, 1} = queryList{queryNum}.title;
        outputData{queryNum+1, 2} = queryList{queryNum}.description;
        outputData{queryNum+1, 3} = table2array(queryList{queryNum}.data);
    end
else % if an existing spreadsheet exists, load it, then append data. 
    % NOTE: Append newest data into column 3, shift old columns to the
    % right
    existingData = readcell('Dupuytrens KPI.xlsx','Sheet','phaseAggregates');
    edHeight = height(existingData);
    edWidth = width(existingData);
    outputData = cell(height(existingData), width(existingData) + 1);
    outputData(:,1:2) = existingData(:,1:2); % paste titles and descriptions from existingData to outputData
    outputData(:,4:end) = existingData(:,3:end); % paste old data from existingData to outputData
    outputData{1,3} = string(datetime('now','Format','yyyy/MM/dd, HH:mm')); % insert current datetime at top of column 3
    for queryNum = 1:length(queryList)
        % check to see if an existing title matches the current query.title
        if any(matches(existingData(:,1), queryList{queryNum}.title))
            outputData{matches(existingData(:,1), queryList{queryNum}.title), 3} = table2array(queryList{queryNum}.data);
        else % if no titles in existing data match the current query.title, create a new row for the current query 
            outputData{edHeight+1, 1} = queryList{queryNum}.title;
            outputData{edHeight+1, 2} = queryList{queryNum}.description;
            outputData{edHeight+1, 3} = table2array(queryList{queryNum}.data);
            edHeight = height(outputData);
        end
    end
end

%%
% --- Arrange the data tables into a singular table for placement into a
% spreadsheet to be imported into Power BI ---
% initialize outputDataPowerBI
% outputDataPowerBI = cell(height(outputData), width(outputData) + 1);
outputDataPowerBI = cell(height(outputData) - 1, 3);
% outputDataPowerBI(1,:) = [{'Survey','Period','Status'}, fliplr(cellstr(extract([outputData{1,3:end}], digitsPattern(4) + '/' + digitsPattern(2) + '/' + digitsPattern(2))))];
% outputDataPowerBI(1,:) = {'Survey','Period','Status','Date','Data'};
% initialize Survey
% outputDataPowerBI(2:end,1) = replace(extractBefore(outputData(2:end,1), ':'), ' Questionnaire', '');
outputDataPowerBI(:,1) = replace(extractBefore(outputData(2:end,1), ':'), ' Questionnaire', '');
% initialize Period
% outputDataPowerBI(contains(outputData(:,1), 'Week'), 2) = deal({'This Week'});
% outputDataPowerBI(contains(outputData(:,1), 'Current Phase'), 2) = deal({'Current Phase'});
% outputDataPowerBI(contains(outputData(:,1), 'Ever'), 2) = deal({'All Time'});
outputDataPowerBI(contains(outputData(2:end,1), 'Week'), 2) = deal({'This Week'});
outputDataPowerBI(contains(outputData(2:end,1), 'Current Phase'), 2) = deal({'Current Phase'});
outputDataPowerBI(contains(outputData(2:end,1), 'Ever'), 2) = deal({'All Time'});
% initialize Status
for statusName = {'Sent','Started','Completed'}
    % outputDataPowerBI(contains(outputData(:,1), statusName), 3) = deal(statusName);
    outputDataPowerBI(contains(outputData(2:end,1), statusName), 3) = deal(statusName);
end
% % assign data
% outputDataPowerBI(2:end, 4:end) = fliplr(outputData(2:end, 3:end));
% assign remaining Survey, Period, Status, Date, and Data
outputDataPowerBI = [{'Survey','Period','Status','Date','Data'}; repmat(outputDataPowerBI, width(outputData) - 2, 1), reshape(repmat(fliplr(cellstr(extract([outputData{1,3:end}], digitsPattern(4) + '/' + digitsPattern(2) + '/' + digitsPattern(2)))), height(outputData) - 1, 1), 2*height(outputDataPowerBI), []), reshape(fliplr(outputData(2:end, 3:end)), 2*height(outputDataPowerBI), [])];


%%
writecell(outputDataPowerBI, 'Dupuytrens KPI for Power BI.xlsx', 'Sheet','phaseAggregates');
% writecell(outputData, 'Dupuytrens KPI.xlsx', 'Sheet','phaseAggregates')

%%
% -------------------------------------------------------------------------
% ------ Page Percentage Section ------
clear queryList

try
    dbConnection = database('FORWARD', 'matt', 'ardent-refurbished-knit');
catch ME
    if strcmp(ME.identifier, 'database:database:dataSourceNameNotFound')
        error('dataSourceNameNotFound. To fix, launch PowerShell as admin, execute the following code, then rerun this script:\nAdd-OdbcDsn -Name ''%s'' -DriverName ''SQL Server'' -DsnType ''System'' -Platform ''64-bit'' -SetPropertyValue ''Server=10.0.100.70''', 'FORWARD')
    else
        rethrow(ME);
    end
    dbsource = "powershell;Add-OdbcDsn -Name 'FORWARD' -DriverName 'SQL Server' -DsnType 'System' -Platform '64-bit' -SetPropertyValue 'Server=10.0.100.70'";
    system(dbsource);
end

% % --- Define common default query components ---
% deafaultQueryComponentList.USE = 'USE Forward ';
% deafaultQueryComponentList.SELECT = 'SELECT up.ActivityDate AS ''ActivityDate'', n.zip AS ''PostalCode'' ';
% deafaultQueryComponentList.Select = 'SELECT up.ActivityDate AS ''ActivityDate'', n.zip AS ''PostalCode'' ';

% --- Define default queries ---
defaultQueryList.availablePhases.description = 'The list of unique survey phases and natalog phaseIDs in Dupuytrens';
defaultQueryList.availablePhases.query = [  'USE Forward ' ...
                                            'SELECT nsr.ColumnTitle as ''NatalogField'', DatabasePrefix, questid, StartDate, EndDate ' ...
                                            'FROM Phase ph ' ...
                                            'JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId ' ...
                                            '    AND ph.ProjectId = 46 -- 46 = Dup --' ...
                                            'WHERE ph.DatabasePrefix like ''{SURVEYNAME}%'' -- must have {SURVEYNAME} database prefix --']; % SURVEYNAME = {'Enroll', 'Dup', 'Short'}

defaultQueryList.availableTables.description = 'The list of unique survey phases and pages in Dupuytrens';
defaultQueryList.availableTables.query = [   'USE Dupuytrens ' ...
                                             'SELECT TABLE_NAME ' ...
                                             '    FROM Dupuytrens.INFORMATION_SCHEMA.TABLES ' ...
                                             '    WHERE TABLE_NAME like ''{SURVEYNAME}%_pg%'' -- must have {SURVEYNAME} in table name --']; % SURVEYNAME = {'Enroll', 'Dup', 'Short'}


defaultQueryList.pageHeaders.description = 'The list of page numbers and page titles ALONG WITH the first survey that started using that specific order of page titles and numbers';
defaultQueryList.pageHeaders.query = [   'USE Forward ' ...
                                         'SELECT DatabasePrefix, ''['' + CAST(PageNumber as VARCHAR(3)) + ''] '' + Pageheader as ''Page Header'' ' ...
                                         '    FROM Page p ' ...
                                         '    JOIN Phase ph ON ph.SurveyToolId = p.SurveyToolId -- match SurveyToolIDs between Forward.Phase and Forward.Page --' ...
                                         '    WHERE ph.DatabasePrefix like ''{SURVEYNAME}%'' -- must have {SURVEYNAME} database prefix --' ...
		                                 '        AND ProjectId = 46 -- specify the projectID... --' ...
                                         '    ORDER BY ph.SurveyToolId, DatabasePrefix, p.PageNumber -- order things ahead of time, for convenience --']; % SURVEYNAME = {'Enroll', 'Dup', 'Short'}

defaultQueryList.questionnairesSent.description = 'Participants who should have completed questionnaires (i.e., the number of questionnaires sent out)';
defaultQueryList.questionnairesSent.query = [  'USE Forward ' ...
                                               'SELECT DISTINCT up.UserId AS ''UserID'', up.ActivityDate AS ''ActivityDate'', n.zip AS ''PostalCode'' ' ...
	                                           'FROM UserPhase up -- grab user activity --' ...
	                                           'JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
	                                           'JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
	                                           'JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
	                                           'WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
		                                       '    AND pj.Name = ''Dupuytrens'' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...
		                                       '    AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
                                                                                                                                        % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}

defaultQueryList.pageStarted.description = 'Participants who started the page';
defaultQueryList.pageStarted.query = [   'USE Forward ' ...
                                         'SELECT DISTINCT up.UserId AS ''UserID'', up.ActivityDate AS ''ActivityDate'', n.zip AS ''PostalCode'' ' ...
	                                     'FROM Dupuytrens.dbo.{SURVEYPAGEHEADER} q -- grab info for specific page of specific questionnaire --' ...
                                         'JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
	                                     'JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
	                                     'JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
	                                     'JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
	                                     'WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
		                                 '    AND pj.Name = ''Dupuytrens'' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...
		                                 '    AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
                                                                                                                                  % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
                                                                                                                                  % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}
                                                                                                                                  % NATALOGPHASEID = whichever natalog phase id matches the current survey (e.g., jul24 for Dup87). natalog phase id and survey names can be found in queryList.availablePhases.(string(currentSurvey)).data

defaultQueryList.pageCompleted.description = 'Participants who completed the page';
defaultQueryList.pageCompleted.query = [ 'USE Forward ' ...
                                         'SELECT DISTINCT up.UserId AS ''UserID'', up.ActivityDate AS ''ActivityDate'', n.zip AS ''PostalCode'' ' ...
	                                     'FROM Dupuytrens.dbo.{SURVEYPAGEHEADER} q -- grab info for specific page of specific questionnaire --' ...
                                         'JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
	                                     'JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
	                                     'JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
	                                     'JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
	                                     'WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
		                                 '    AND pj.Name = ''Dupuytrens'' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...
		                                 '    AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --', ...
                                         '    AND {NATALOGPHASEID} = ph.CompleteStatus']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
                                                                                          % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
                                                                                          % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}      
                                                                                          % NATALOGPHASEID = whichever natalog phase id matches the current survey (e.g., jul24 for Dup87). natalog phase id and survey names can be found in queryList.availablePhases.(string(currentSurvey)).data

defaultQueryList.pageUnfinished.description = 'Participants who didn''t complete the page';
defaultQueryList.pageUnfinished.query = [   'USE Forward ' ...
                                            'SELECT DISTINCT up.UserId AS ''UserID'', up.ActivityDate AS ''ActivityDate'', n.zip AS ''PostalCode'' ' ...
	                                        'FROM Dupuytrens.dbo.{SURVEYPAGEHEADER} q -- grab info for specific page of specific questionnaire --' ...
                                            'JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
	                                        'JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
	                                        'JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
	                                        'JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
	                                        'WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
		                                    '    AND pj.Name = ''Dupuytrens'' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...
		                                    '    AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --'...
                                            '    AND {NATALOGPHASEID} = ph.StartedStatus']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
                                                                                            % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
                                                                                            % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}
                                                                                            % NATALOGPHASEID = whichever natalog phase id matches the current survey (e.g., jul24 for Dup87). natalog phase id and survey names can be found in queryList.availablePhases.(string(currentSurvey)).data

%%
tic
try
    % --- Determine the list of page numbers and page titles ALONG WITH the
    % first survey that started using that specific order of page titles
    % and numbers for Enrollment, Long Surveys, and Short Surveys ---
    % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'Enrollment', 'monthly_YYYY03MM', 'followup_%'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug', 'followup_pg8' NOT 'followup_6_mo_pg8'))
    % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''monthly_YYYY03MM''', 'ph.DatabasePrefix = ''followup_20250107''' OR 'ph.DatabasePrefix = ''followup_6_mo'''} 
    % SURVEYDESCRIPTION = {'Psoriasis Registry - Enrollment PhaseId', 'Psoriasis Monthly Survey for MMM YYYY', 'Psoriasis Followup Survey for version "20250107" or "6_mo"'}
    surveyList = {'Enrollment', 'Long', 'Short'};
    for currentSurvey = surveyList
        switch string(currentSurvey)
            case 'Enrollment'
                surveyName = 'Enroll';
                surveySelect = 'ph.Name LIKE ''%Enrollment%''';
                surveyDescription = 'Dupuytrens Enrollment PhaseId';
                surveyPageHeader = [sprintf('%s_pg', 'EnrollDup'), '%d'];
            case 'Long'
                surveyName = 'Dup';
                % surveySelect = 'ph.DatabasePrefix = ''Dup%d''';
                % surveyDescription = 'Dupuytrens Long Survey for ph%d';
                % surveyPageHeader = [sprintf('%s_pg', 'Dup%d'), '%d'];
                surveySelect = 'ph.DatabasePrefix = ''Dup%s''';
                surveyDescription = 'Dupuytrens Long Survey for ph%s';
                surveyPageHeader = [sprintf('%s_pg', 'Dup%s'), '%d'];
            case 'Short'
                surveyName = 'Short';
                % surveySelect = 'ph.DatabasePrefix = ''Short%d''';
                % surveyDescription = 'Dupuytrens Short Survey for ph%d';
                % surveyPageHeader = [sprintf('%s_pg', 'Short%d'), '%d'];
                surveySelect = 'ph.DatabasePrefix = ''Short%s''';
                surveyDescription = 'Dupuytrens Short Survey for ph%s';
                surveyPageHeader = [sprintf('%s_pg', 'Short%s'), '%d'];
        end
        
        % initialize queryList structs
        [queryList.pageHeaders.(string(currentSurvey)), ...
           queryList.availablePhases.(string(currentSurvey)), ...
           queryList.availableTables.(string(currentSurvey)), ...
           queryList.questionnairesSent.(string(currentSurvey)), ...
           queryList.pageStarted.(string(currentSurvey)), ...
           queryList.pageCompleted.(string(currentSurvey)), ...
           queryList.pageUnfinished.(string(currentSurvey))] = deal(struct);

        % query pageHeaders
        queryList.pageHeaders.(string(currentSurvey)) = evalQuery(queryList.pageHeaders.(string(currentSurvey)), ...
                                                                  sprintf('%s Questionnaire: PageHeaders', string(currentSurvey)), ...
                                                                  defaultQueryList.pageHeaders.description, ...
                                                                  strrep(defaultQueryList.pageHeaders.query, '{SURVEYNAME}', surveyName), ...
                                                                  dbConnection);

        % - Determine the list of pages that can be queried for each survey -
        % for each database prefix, the initial order of page headers is
        % screwy (e.g., [1], [10], [11], ..., [19], [2], [20], ...). Fix
        % this. -

        % create another column in ~.data that is a two-digit representation 
        % of the page number. The rows of ~.data will be sorted by this 
        % two-digit number
        queryList.pageHeaders.(string(currentSurvey)).data.FormattedPageNum = extractBetween(queryList.pageHeaders.(string(currentSurvey)).data.PageHeader, '[', ']');
        queryList.pageHeaders.(string(currentSurvey)).data.FormattedPageNum(matches(queryList.pageHeaders.(string(currentSurvey)).data.FormattedPageNum, digitsPattern(1))) = cellstr([repmat(['0'], sum(matches(queryList.pageHeaders.(string(currentSurvey)).data.FormattedPageNum, digitsPattern(1))), 1), [queryList.pageHeaders.(string(currentSurvey)).data.FormattedPageNum{matches(queryList.pageHeaders.(string(currentSurvey)).data.FormattedPageNum, digitsPattern(1))}]']);
        queryList.pageHeaders.(string(currentSurvey)).data = sortrows(queryList.pageHeaders.(string(currentSurvey)).data, {'DatabasePrefix', 'FormattedPageNum'});
        
        % create a list of database prefix numbers AND a list page headers 
        % for each database prefix
        databasePrefixNumList.(string(currentSurvey)) = [];
        databasePrefixList.(string(currentSurvey)) = {};
        % for databasePrefixName = unique(queryList.pageHeaders.(string(currentSurvey)).data.DatabasePrefix)'
        for databasePrefixName = unique(queryList.pageHeaders.(string(currentSurvey)).data.DatabasePrefix)'
            databasePrefixNumList.(string(currentSurvey)) = [databasePrefixNumList.(string(currentSurvey)), str2double(extract(databasePrefixName, digitsPattern))];
            % databasePrefixList.(string(currentSurvey)) = [databasePrefixList.(string(currentSurvey)), extractAfter(databasePrefixName, [surveyName, '_'])];
            databasePrefixList.(string(currentSurvey)) = [databasePrefixList.(string(currentSurvey)), extractAfter(databasePrefixName, surveyName)];
            pageHeaders.(string(currentSurvey)).(string(databasePrefixName)) = queryList.pageHeaders.(string(currentSurvey)).data.PageHeader(matches(queryList.pageHeaders.(string(currentSurvey)).data.DatabasePrefix, databasePrefixName));
        end
        
        if ~strcmp(currentSurvey, 'Enrollment')
            % - Determine the list of available phases and corresponding Natalog phase ids for the current survey -
            queryList.availablePhases.(string(currentSurvey)) = evalQuery(queryList.availablePhases.(string(currentSurvey)), ...
                                                                          sprintf('%s Questionnaire: Available Phases', string(currentSurvey)), ...
                                                                          defaultQueryList.availablePhases.description, ...
                                                                          strrep(defaultQueryList.availablePhases.query, '{SURVEYNAME}', surveyName), ...
                                                                          dbConnection);

            % replace empty questids with 'NULL'
            queryList.availablePhases.(string(currentSurvey)).data.questid(cellfun(@isempty, queryList.availablePhases.(string(currentSurvey)).data.questid)) = deal({'NULL'});
            
            % sort the pageHeaders by date (really it's just alphabetically)
            [pageHeaders.(string(currentSurvey)), sortIndex] = orderfields(pageHeaders.(string(currentSurvey)));


            % retreive list of available table names (e.g., Dup86_pg7, etc)
            queryList.availableTables.(string(currentSurvey)) = evalQuery(queryList.availableTables.(string(currentSurvey)), ...
                                                                          sprintf('%s Questionnaire: Available Tables', string(currentSurvey)), ...
                                                                          defaultQueryList.availableTables.description, ...
                                                                          strrep(defaultQueryList.availableTables.query, '{SURVEYNAME}', surveyName), ...
                                                                          dbConnection);
            
            for phaseName = databasePrefixList.(string(currentSurvey)) % for each survey phase name...
                % define current survey description
                currentSurveyDescription = sprintf(surveyDescription, string(phaseName));

                % identify the Natalog Phase ID that corresponds to the
                % current survey phase
                currentNatalogPhaseID = queryList.availablePhases.(string(currentSurvey)).data.NatalogField{matches(queryList.availablePhases.(string(currentSurvey)).data.DatabasePrefix, sprintf('%s%s', surveyName, string(phaseName)))};
                
                % - Total questionnaires sent -
                queryList.questionnairesSent.(string(currentSurvey)) = evalQuery(queryList.questionnairesSent.(string(currentSurvey)), ...
                                                                                 sprintf('%s Questionnaire: Sent Surveys', string(currentSurvey)), ...
                                                                                 defaultQueryList.questionnairesSent.description, ...
                                                                                 sprintf(strrep(strrep(defaultQueryList.questionnairesSent.query, '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', currentSurveyDescription), string(phaseName)), ...
                                                                                 dbConnection, ...
                                                                                 sprintf('%s%s', surveyName, string(phaseName)));


                for pageNum = str2double(extractBetween(pageHeaders.(string(currentSurvey)).(sprintf('%s%s', surveyName, string(phaseName))), '[', ']'))'
                    currentSurveyPageHeader = sprintf(surveyPageHeader, string(phaseName), pageNum);

                    % - Total patients who completed, started, or didn't finish the current page -
                    if ~any(matches(queryList.availableTables.(string(currentSurvey)).data.TABLE_NAME, sprintf('%s%s_pg%d', surveyName, string(phaseName), pageNum))) % if current table name can't be found in Psoriasis tables, then the current page must be pgDrug

                        queryList.pageStarted.(string(currentSurvey)) = evalQuery(queryList.pageStarted.(string(currentSurvey)), ...
                                                                                  sprintf('%s Questionnaire: Started Surveys', string(currentSurvey)), ...
                                                                                  defaultQueryList.pageStarted.description, ...
                                                                                  sprintf(strrep(strrep(strrep(strrep(strrep(defaultQueryList.pageStarted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', currentSurveyDescription), 'pg%d', 'pgDrug'), string(phaseName), string(phaseName), string(phaseName)), ...
                                                                                  dbConnection, ...
                                                                                  sprintf('%s%s', surveyName, string(phaseName)), ...
                                                                                  sprintf('pg%d', pageNum));
                        queryList.pageCompleted.(string(currentSurvey)) = evalQuery(queryList.pageCompleted.(string(currentSurvey)), ...
                                                                                    sprintf('%s Questionnaire: Completed Surveys', string(currentSurvey)), ...
                                                                                    defaultQueryList.pageCompleted.description, ...
                                                                                    sprintf(strrep(strrep(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', currentSurveyDescription), 'pg%d', 'pgDrug'), string(phaseName), string(phaseName), string(phaseName)), ...
                                                                                    dbConnection, ...
                                                                                    sprintf('%s%s', surveyName, string(phaseName)), ...
                                                                                    sprintf('pg%d', pageNum));
                        queryList.pageUnfinished.(string(currentSurvey)) = evalQuery(queryList.pageUnfinished.(string(currentSurvey)), ...
                                                                                     sprintf('%s Questionnaire: Unfinished Surveys', string(currentSurvey)), ...
                                                                                     defaultQueryList.pageUnfinished.description, ...
                                                                                     sprintf(strrep(strrep(strrep(strrep(strrep(defaultQueryList.pageUnfinished.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', currentSurveyDescription), 'pg%d', 'pgDrug'), string(phaseName), string(phaseName), string(phaseName)), ...
                                                                                     dbConnection, ...
                                                                                     sprintf('%s%s', surveyName, string(phaseName)), ...
                                                                                     sprintf('pg%d', pageNum));

                    else
                        queryList.pageStarted.(string(currentSurvey)) = evalQuery(queryList.pageStarted.(string(currentSurvey)), ...
                                                                                  sprintf('%s Questionnaire: Started Surveys', string(currentSurvey)), ...
                                                                                  defaultQueryList.pageStarted.description, ...
                                                                                  sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageStarted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', currentSurveyDescription), string(phaseName), pageNum, string(phaseName), pageNum, string(phaseName)), ...
                                                                                  dbConnection, ...
                                                                                  sprintf('%s%s', surveyName, string(phaseName)), ...
                                                                                  sprintf('pg%d', pageNum));
                        queryList.pageCompleted.(string(currentSurvey)) = evalQuery(queryList.pageCompleted.(string(currentSurvey)), ...
                                                                                    sprintf('%s Questionnaire: Completed Surveys', string(currentSurvey)), ...
                                                                                    defaultQueryList.pageCompleted.description, ...
                                                                                    sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', currentSurveyDescription), string(phaseName), pageNum, string(phaseName), pageNum, string(phaseName)), ...
                                                                                    dbConnection, ...
                                                                                    sprintf('%s%s', surveyName, string(phaseName)), ...
                                                                                    sprintf('pg%d', pageNum));
                        queryList.pageUnfinished.(string(currentSurvey)) = evalQuery(queryList.pageUnfinished.(string(currentSurvey)), ...
                                                                                     sprintf('%s Questionnaire: pageUnfinished Surveys', string(currentSurvey)), ...
                                                                                     defaultQueryList.pageUnfinished.description, ...
                                                                                     sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageUnfinished.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', currentSurveyDescription), string(phaseName), pageNum, string(phaseName), pageNum, string(phaseName)), ...
                                                                                     dbConnection, ...
                                                                                     sprintf('%s%s', surveyName, string(phaseName)), ...
                                                                                     sprintf('pg%d', pageNum));
                    end
                    % % pageUnfinished = pageStarted - pageCompleted
                    % queryList.pageUnfinished.(string(currentSurvey)).title = sprintf('%s Questionnaire: Unfinished Surveys', string(currentSurvey));
                    % queryList.pageUnfinished.(string(currentSurvey)).description = 'Participants who started, but didn''t complete the page';
                    % % queryList.pageUnfinished.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum)) = array2table(table2array(queryList.pageStarted.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum))), 'VariableNames',{'NumParticipantsThatNoFinishPage'});
                    % % queryList.pageUnfinished.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum)) = array2table(table2array(queryList.pageStarted.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum))), 'VariableNames',queryList.pageStarted.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum)).Properties.VariableNames);
                    % queryList.pageUnfinished.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum)) = queryList.pageStarted.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum))(~matches(queryList.pageStarted.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum)).UserID, queryList.pageCompleted.(string(currentSurvey)).data.(sprintf('%s%s', surveyName, string(phaseName))).(sprintf('pg%d', pageNum)).UserID), :); % extract the entries in pageStarted that aren't found in pageCompleted (i.e., unfinished pages)               

                    fprintf('%s Survey, Phase %s, Page %d completed...\n', string(currentSurvey), string(phaseName), pageNum);
                end
            end
        else
            % - Determine the list of available phases and corresponding Natalog phase ids for the current survey -
            queryList.availablePhases.(string(currentSurvey)) = evalQuery(queryList.availablePhases.(string(currentSurvey)), ...
                                                                          sprintf('%s Questionnaire: Available Phases', string(currentSurvey)), ...
                                                                          defaultQueryList.availablePhases.description, ...
                                                                          strrep(defaultQueryList.availablePhases.query, '{SURVEYNAME}', surveyName), ...
                                                                          dbConnection);

            % replace empty questids with 'NULL'
            queryList.availablePhases.(string(currentSurvey)).data.questid(cellfun(@isempty, queryList.availablePhases.(string(currentSurvey)).data.questid)) = deal({'NULL'});

            % - Total questionnaires sent -
            queryList.questionnairesSent.(string(currentSurvey)) = evalQuery(queryList.questionnairesSent.(string(currentSurvey)), ...
                                                                             sprintf('%s Questionnaire: Sent Surveys', string(currentSurvey)), ...
                                                                             defaultQueryList.questionnairesSent.description, ...
                                                                             strrep(strrep(defaultQueryList.questionnairesSent.query, '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), ...
                                                                             dbConnection, ...
                                                                             extractBefore(surveyPageHeader, '_'));

            % retreive list of available table names (e.g., PSOR86_pg7, etc)
            queryList.availableTables.(string(currentSurvey)) = evalQuery(queryList.availableTables.(string(currentSurvey)), ...
                                                                          sprintf('%s Questionnaire: Available Tables', string(currentSurvey)), ...
                                                                          defaultQueryList.availableTables.description, ...
                                                                          strrep(defaultQueryList.availableTables.query, '{SURVEYNAME}', surveyName), ...
                                                                          dbConnection);

            % identify the Natalog Phase ID that corresponds to the
            % current survey phase
            currentNatalogPhaseID = queryList.availablePhases.(string(currentSurvey)).data.NatalogField{matches(queryList.availablePhases.(string(currentSurvey)).data.DatabasePrefix, databasePrefixName)};
            
            for pageNum = str2double(extractBetween(pageHeaders.(string(currentSurvey)).(extractBefore(surveyPageHeader, '_')), '[', ']'))'
                % - Total patients who completed, started, or didn't finish the current page -
                if ~any(matches(queryList.availableTables.(string(currentSurvey)).data.TABLE_NAME, sprintf('%s_pg%d', extractBefore(surveyPageHeader, '_'), pageNum))) % if current table name can't be found in Psoriasis tables, then the current page must be pgDrug
                    queryList.pageStarted.(string(currentSurvey)) = evalQuery(queryList.pageStarted.(string(currentSurvey)), ...
                                                                              sprintf('%s Questionnaire: Started Surveys', string(currentSurvey)), ...
                                                                              defaultQueryList.pageStarted.description, ...
                                                                              strrep(strrep(strrep(strrep(strrep(defaultQueryList.pageStarted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), 'pg%d', 'pgDrug'), ...
                                                                              dbConnection, ...
                                                                              extractBefore(surveyPageHeader, '_'), ...
                                                                              sprintf('pg%d', pageNum));
                    queryList.pageCompleted.(string(currentSurvey)) = evalQuery(queryList.pageCompleted.(string(currentSurvey)), ...
                                                                                sprintf('%s Questionnaire: Completed Surveys', string(currentSurvey)), ...
                                                                                defaultQueryList.pageCompleted.description, ...
                                                                                strrep(strrep(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), 'pg%d', 'pgDrug'), ...
                                                                                dbConnection, ...
                                                                                extractBefore(surveyPageHeader, '_'), ...
                                                                                sprintf('pg%d', pageNum));
                    queryList.pageUnfinished.(string(currentSurvey)) = evalQuery(queryList.pageUnfinished.(string(currentSurvey)), ...
                                                                                 sprintf('%s Questionnaire: pageUnfinished Surveys', string(currentSurvey)), ...
                                                                                 defaultQueryList.pageUnfinished.description, ...
                                                                                 strrep(strrep(strrep(strrep(strrep(defaultQueryList.pageUnfinished.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), 'pg%d', 'pgDrug'), ...
                                                                                 dbConnection, ...
                                                                                 extractBefore(surveyPageHeader, '_'), ...
                                                                                 sprintf('pg%d', pageNum));
                else
                    queryList.pageStarted.(string(currentSurvey)) = evalQuery(queryList.pageStarted.(string(currentSurvey)), ...
                                                                              sprintf('%s Questionnaire: Started Surveys', string(currentSurvey)), ...
                                                                              defaultQueryList.pageStarted.description, ...
                                                                              sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageStarted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', strrep(surveySelect, '%', '%%')), '{SURVEYDESCRIPTION}', surveyDescription), pageNum, pageNum), ...
                                                                              dbConnection, ...
                                                                              extractBefore(surveyPageHeader, '_'), ...
                                                                              sprintf('pg%d', pageNum));
                    queryList.pageCompleted.(string(currentSurvey)) = evalQuery(queryList.pageCompleted.(string(currentSurvey)), ...
                                                                                sprintf('%s Questionnaire: Completed Surveys', string(currentSurvey)), ...
                                                                                defaultQueryList.pageCompleted.description, ...
                                                                                sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', strrep(surveySelect, '%', '%%')), '{SURVEYDESCRIPTION}', surveyDescription), pageNum, pageNum), ...
                                                                                dbConnection, ...
                                                                                extractBefore(surveyPageHeader, '_'), ...
                                                                                sprintf('pg%d', pageNum));
                    queryList.pageUnfinished.(string(currentSurvey)) = evalQuery(queryList.pageUnfinished.(string(currentSurvey)), ...
                                                                                 sprintf('%s Questionnaire: pageUnfinished Surveys', string(currentSurvey)), ...
                                                                                 defaultQueryList.pageUnfinished.description, ...
                                                                                 sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageUnfinished.query, '{NATALOGPHASEID}', currentNatalogPhaseID), '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', strrep(surveySelect, '%', '%%')), '{SURVEYDESCRIPTION}', surveyDescription), pageNum, pageNum), ...
                                                                                 dbConnection, ...
                                                                                 extractBefore(surveyPageHeader, '_'), ...
                                                                                 sprintf('pg%d', pageNum));
                end
                % % pageUnfinished = pageStarted - pageCompleted
                % queryList.pageUnfinished.(string(currentSurvey)).title = sprintf('%s Questionnaire: Unfinished Surveys', string(currentSurvey));
                % queryList.pageUnfinished.(string(currentSurvey)).description = 'Participants who started, but didn''t complete the page';
                % % queryList.pageUnfinished.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum)) = array2table(table2array(queryList.pageStarted.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum))), 'VariableNames',{'NumParticipantsThatNoFinishPage'});
                % % queryList.pageUnfinished.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum)) = array2table(table2array(queryList.pageStarted.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum))), 'VariableNames',queryList.pageStarted.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum)).Properties.VariableNames);
                % queryList.pageUnfinished.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum)) = queryList.pageStarted.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum))(~matches(queryList.pageStarted.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum)).UserID, queryList.pageCompleted.(string(currentSurvey)).data.(string(extractBefore(surveyPageHeader, '_'))).(sprintf('pg%d', pageNum)).UserID), :); % extract the entries in pageStarted that aren't found in pageCompleted (i.e., unfinished pages)

                fprintf('%s Survey, Page %d completed...\n', string(currentSurvey), pageNum);
            end
        end

    end

catch ME
    close(dbConnection);
    rethrow(ME);
end
queryListFieldNames = fieldnames(queryList);
toc

% save current queryList (in case you want the original queryList without
% rerunning all the queries)
queryListRaw = queryList;

% --- Replace each postal code (if a US Zip Code) with the state abbreviation ---
% load zip code key
warning off
zipRawTable = readtable('ZIP_Locale_Detail.xls');
warning on
% % find list of unique 3-digit zips and their corresponding states
uniqueStateList = [unique(zipRawTable.PHYSICALSTATE); {'International'}];
%%
% replace postal codes
tic
fprintf('Replacing postal codes with state abbreviations...\n')
for currentSurvey = surveyList
    for tableType = {'questionnairesSent', 'pageStarted', 'pageCompleted', 'pageUnfinished'}
        for currentPhaseName = fieldnames(queryList.(string(tableType)).(string(currentSurvey)).data)'
            switch string(tableType)
                case 'questionnairesSent'
                    % replace USA postal codes with the corresponding state
                    % abbreviation
                    queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).PostalCode = replace(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).PostalCode, zipRawTable.DELIVERYZIPCODE, zipRawTable.PHYSICALSTATE);
                    % replace postal codes that weren't found to be US states
                    % with 'International'
                    queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).PostalCode(~matches(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).PostalCode, lettersPattern(2))) = {'International'};
                otherwise
                    for pageName = fieldnames(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)))'
                        % replace USA postal codes with the corresponding state
                        % abbreviation
                        queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).(string(pageName)).PostalCode = replace(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).(string(pageName)).PostalCode, zipRawTable.DELIVERYZIPCODE, zipRawTable.PHYSICALSTATE);
                        % replace postal codes that weren't found to be US states
                        % with 'International'
                        queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).(string(pageName)).PostalCode(~matches(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseName)).(string(pageName)).PostalCode, lettersPattern(2))) = {'International'};
                    end
            end
        end
    end
end
toc

%%
% --- Arrange the data into a neat table ---
tic
fprintf('Arranging data into a neat table...\n')

tableTypeList = {'percentPageCompleted', 'pageCompleted', 'percentPageStarted', 'pageStarted', 'percentPageUnfinished', 'pageUnfinished'};
variationTypeList = {'Standard','Weekday','PostalState'};
weekdayList = {'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'};
for currentSurvey = surveyList
    % each pageHeader will have its own table (cus each pageHeader has
    % different pages). determine which phases go with which
    % pageHeaders
    % determine the unique pageHeader sets
    uniqueHeadersCount = 1;
    phaseList.(string(currentSurvey)) = flipud(fieldnames(pageHeaders.(string(currentSurvey))))'; % NOTE: the order of the fieldnames has been reversed to (hopefully) facilitate the formatting onto the spreadsheet at the end (e.g., where the most recent surveys start at the top of the spreadsheet)
    pageHeadersUnique.(string(currentSurvey)).headers{uniqueHeadersCount} = pageHeaders.(string(currentSurvey)).(string(phaseList.(string(currentSurvey))(1))); % initialize the first set of unique pageHeaders as the set that belongs to the first phase
    pageHeadersUnique.(string(currentSurvey)).phases{uniqueHeadersCount} = phaseList.(string(currentSurvey))(1); % initialize the first phase of the first set of unique pageHeaders as the first phase
    for phaseListNum = 2:length(phaseList.(string(currentSurvey)))
        % if the headers match between the current set unique
        % pageHeaders and the pageHeaders for the current phase AND the
        % current set of unique pageHeaders and the pageHeaders for the
        % current phase have the same number of elements...
        if all(matches(pageHeadersUnique.(string(currentSurvey)).headers{uniqueHeadersCount}, pageHeaders.(string(currentSurvey)).(string(phaseList.(string(currentSurvey))(phaseListNum))))) && (numel(pageHeadersUnique.(string(currentSurvey)).headers{uniqueHeadersCount}) == numel(pageHeaders.(string(currentSurvey)).(string(phaseList.(string(currentSurvey))(phaseListNum)))))
            % add the current phase to the list of phases that share
            % the current set of unique pageHeaders
            pageHeadersUnique.(string(currentSurvey)).phases{uniqueHeadersCount} = [pageHeadersUnique.(string(currentSurvey)).phases{uniqueHeadersCount}, phaseList.(string(currentSurvey))(phaseListNum)];
        else % if the current set of unique pageHeaders and the pageHeaders for the current phase are different...
            uniqueHeadersCount = uniqueHeadersCount + 1;
            pageHeadersUnique.(string(currentSurvey)).headers{uniqueHeadersCount} = pageHeaders.(string(currentSurvey)).(string(phaseList.(string(currentSurvey))(phaseListNum))); % define the next set of unique pageHeaders as the set that belongs to the current phase
            pageHeadersUnique.(string(currentSurvey)).phases{uniqueHeadersCount} = phaseList.(string(currentSurvey))(phaseListNum); % initialize the first phase of the next set of unique pageHeaders as the current phase
        end
    end

    for pageHeaderUniqueNum = 1:length(pageHeadersUnique.(string(currentSurvey)).headers)
        % - Initialize outputDataTablesCell tables for current pageHeader -
        for tableType = {'pageStarted', 'pageCompleted', 'pageUnfinished', 'percentPageStarted', 'percentPageCompleted', 'percentPageUnfinished'}
            for variationType = variationTypeList
                % initialize outputDataTablesCell tables for current pageHeader
                 switch string(variationType)
                    case 'Standard'
                        outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum} = cell(length(pageHeadersUnique.(string(currentSurvey)).phases{pageHeaderUniqueNum}) + 1, height(pageHeadersUnique.(string(currentSurvey)).headers{pageHeaderUniqueNum}) + 1);
                        outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}(2:end, 1) = pageHeadersUnique.(string(currentSurvey)).phases{pageHeaderUniqueNum}';
                        outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}(1, 2:end) = pageHeadersUnique.(string(currentSurvey)).headers{pageHeaderUniqueNum}';
                    case 'Weekday'
                        outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum} = cell((length(weekdayList) * length(pageHeadersUnique.(string(currentSurvey)).phases{pageHeaderUniqueNum})) + 1, height(pageHeadersUnique.(string(currentSurvey)).headers{pageHeaderUniqueNum}) + 1);
                        outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(2:end, 1) = reshape(repmat(pageHeadersUnique.(string(currentSurvey)).phases{pageHeaderUniqueNum}, length(weekdayList), 1), [], 1);
                        outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(1, 2:end) = pageHeadersUnique.(string(currentSurvey)).headers{pageHeaderUniqueNum}';
                    case 'PostalState'
                        outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum} = cell((length(uniqueStateList) * length(pageHeadersUnique.(string(currentSurvey)).phases{pageHeaderUniqueNum})) + 1, height(pageHeadersUnique.(string(currentSurvey)).headers{pageHeaderUniqueNum}) + 1);
                        outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(2:end, 1) = reshape(repmat(pageHeadersUnique.(string(currentSurvey)).phases{pageHeaderUniqueNum}, length(uniqueStateList), 1), [], 1);
                        outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(1, 2:end) = pageHeadersUnique.(string(currentSurvey)).headers{pageHeaderUniqueNum}';
                end
                
                % insert data for each page for the current pageHeader
                switch string(variationType)
                    case 'Standard'
                        currentPhaseNameList = unique(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}(2:end, 1)', 'stable');
                    otherwise
                        currentPhaseNameList = unique(extractBefore(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}(2:end, 1)', '/'), 'stable');
                end
                pageNumList = str2double(extractBetween(pageHeadersUnique.(string(currentSurvey)).headers{pageHeaderUniqueNum}, '[', ']'))'; % assume all tables have the same page numbers as pageStarted
                for currentPhaseNameIndex = 1:length(currentPhaseNameList)
                    for pageNumIndex = 1:length(pageNumList)
                        warning off
                        if contains(tableType, 'percent')
                            switch string(variationType)
                                case 'Standard'
                                    % - total percents of pages -
                                    switch string(tableType)
                                        case 'percentPageStarted' % should be percent of {surveys started}/{total surveys sent}
                                            outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}{currentPhaseNameIndex + 1, pageNumIndex + 1} = round(10000 * (height(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex)))) / height(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex)))))) / 100;
                                        otherwise % should be a percent of {surveys either completed or unfinished}/{total surveys started}
                                            outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}{currentPhaseNameIndex + 1, pageNumIndex + 1} = round(10000 * (height(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex)))) / height(queryList.pageStarted.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex)))))) / 100;
                                    end
                                case 'Weekday'
                                    % - weekdays -
                                    % assign the list of weekday names to the first column
                                    outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 1 + [1:length(weekdayList)], 1) = weekdayList';
                                    % evaluate the percent of surveys that 
                                    % were started, completed, or unfinished 
                                    % on each weekday relative to the total
                                    % number of surveys for the week (i.e.,
                                    % {total surveys with a given status for
                                    % a weekday}/{sum of total surveys with
                                    % a given status for the whole week})
                                    outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 1 + [1:length(weekdayList)], pageNumIndex + 1) = num2cell(round((10000 * circshift(sum(weekday(extract(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).ActivityDate, digitsPattern(4) + '-' + digitsPattern(2) + '-' + digitsPattern(2))) == [1:length(weekdayList)], 1, 'omitnan')', -1)) / height(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))))) / 100);
                                    % switch string(tableType)
                                    %     case 'percentPageStarted' % should be percent of {surveys started}/{total surveys sent}
                                    %         outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 1 + [1:length(weekdayList)], pageNumIndex + 1) = num2cell(round((10000 * circshift(sum(weekday(extract(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).ActivityDate, digitsPattern(4) + '-' + digitsPattern(2) + '-' + digitsPattern(2))) == [1:length(weekdayList)], 1, 'omitnan')', -1)) / height(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))))) / 100);
                                    %     otherwise % should be a percent of {surveys either completed or unfinished}/{total surveys started}
                                    %         outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 1 + [1:length(weekdayList)], pageNumIndex + 1) = num2cell(round((10000 * circshift(sum(weekday(extract(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).ActivityDate, digitsPattern(4) + '-' + digitsPattern(2) + '-' + digitsPattern(2))) == [1:length(weekdayList)], 1, 'omitnan')', -1)) / height(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))))) / 100);
                                    % end
                                case 'PostalState'
                                    % - postal codes -
                                    % assign the list of unique postal code state names to the first column
                                    outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(uniqueStateList)*(currentPhaseNameIndex-1) + 1 + [1:length(uniqueStateList)], 1) = uniqueStateList;
                                    % evaluate the percent counts of postal codes for each page
                                    switch string(tableType)
                                        case 'percentPageStarted' % should be percent of {surveys started by a state}/{total surveys sent to that state}
                                            outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(uniqueStateList)*(currentPhaseNameIndex-1) + 1 + [1:length(uniqueStateList)], pageNumIndex + 1) = num2cell(round((10000 * sum(string(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).PostalCode) == uniqueStateList', 1, 'omitnan')') ./ sum(string(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).PostalCode) == uniqueStateList', 1, 'omitnan')') / 100);
                                        otherwise % should be a percent of {surveys either completed or unfinished by a state}/{total surveys started by that state}
                                            outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(uniqueStateList)*(currentPhaseNameIndex-1) + 1 + [1:length(uniqueStateList)], pageNumIndex + 1) = num2cell(round((10000 * sum(string(queryList.(strrep(string(tableType), 'ercentP', '')).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).PostalCode) == uniqueStateList', 1, 'omitnan')') ./ sum(string(queryList.pageStarted.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).PostalCode) == uniqueStateList', 1, 'omitnan')') / 100);
                                    end
                            end
                        else
                            switch string(variationType)
                                case 'Standard'
                                    % - total number of pages -
                                    outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}(currentPhaseNameIndex + 1, pageNumIndex + 1) = {height(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))))};
                                case 'Weekday'
                                    % - weekdays -
                                    % assign the list of weekday names to the first column
                                    outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 1 + [1:length(weekdayList)], 1) = weekdayList';
                                    % evaluate the weekday that each page was active 
                                    outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 1 + [1:length(weekdayList)], pageNumIndex + 1) = num2cell(circshift(sum(weekday(extract(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).ActivityDate, digitsPattern(4) + '-' + digitsPattern(2) + '-' + digitsPattern(2))) == [1:length(weekdayList)], 1, 'omitnan')', -1));
                                case 'PostalState'
                                    % - postal codes -
                                    % assign the list of unique postal code state names to the first column
                                    outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(uniqueStateList)*(currentPhaseNameIndex-1) + 1 + [1:length(uniqueStateList)], 1) = uniqueStateList;
                                    % evaluate the counts of postal codes for each page
                                    outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(uniqueStateList)*(currentPhaseNameIndex-1) + 1 + [1:length(uniqueStateList)], pageNumIndex + 1) = num2cell(sum(string(queryList.(string(tableType)).(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))).PostalCode) == uniqueStateList', 1, 'omitnan')');
                            end
                        end
                        warning on
                    end
                    
                    % append number of questionnaires sent as (n=####) to phase 
                    % name in first column. For tables that have text in the 
                    % first column, append the existing text to the (n=####) info
                    currentQuestID = queryList.availablePhases.(string(currentSurvey)).data.questid{matches(queryList.availablePhases.(string(currentSurvey)).data.DatabasePrefix, currentPhaseNameList{currentPhaseNameIndex})};
                    switch string(variationType)
                        case 'Standard'
                            outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}(currentPhaseNameIndex + 1, 1) = {sprintf('%s/%s (n=%d)', currentPhaseNameList{currentPhaseNameIndex}, currentQuestID, height(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex)))))};
                        case 'Weekday'
                            outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 2 : length(weekdayList)*currentPhaseNameIndex + 1, 1) = join([repmat({sprintf('%s/%s (n=%d)', currentPhaseNameList{currentPhaseNameIndex}, currentQuestID, height(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex)))))}, length(weekdayList), 1), outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(weekdayList)*(currentPhaseNameIndex-1) + 2 : length(weekdayList)*currentPhaseNameIndex + 1, 1)], ' - ');
                        case 'PostalState'
                            outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(uniqueStateList)*(currentPhaseNameIndex-1) + 2 : length(uniqueStateList)*currentPhaseNameIndex + 1, 1) = join([repmat({sprintf('%s/%s (n=%d)', currentPhaseNameList{currentPhaseNameIndex}, currentQuestID, height(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex)))))}, length(uniqueStateList), 1), outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(length(uniqueStateList)*(currentPhaseNameIndex-1) + 2 : length(uniqueStateList)*currentPhaseNameIndex + 1, 1)], ' - ');
                    end
                end

                % % replace any NaN values with 0
                % switch string(variationType)
                %     case 'Standard'
                %         outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}(cellfun(@any, cellfun(@isnan, outputDataTablesCell.(string(tableType)).(string(currentSurvey)){pageHeaderUniqueNum}, 'UniformOutput', false))) = {0};
                %     otherwise
                %         outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}(cellfun(@any, cellfun(@isnan, outputDataTablesCell.(sprintf('%s%s', string(tableType), string(variationType))).(string(currentSurvey)){pageHeaderUniqueNum}, 'UniformOutput', false))) = {0};
                % end
            end
        end
    end
    fprintf('Arranged %s data...\n', string(currentSurvey));
end
toc

%%
% --- Arrange the data tables into a singular table for placement into a
% spreadsheet to be imported into Power BI ---

for variationType = variationTypeList
    % identify modified tableTypeList names
    switch string(variationType)
        case 'Standard'
            tableTypeListMod = tableTypeList;
        otherwise
            tableTypeListMod = join([tableTypeList', repmat(variationType, length(tableTypeList), 1)], '')';
    end
    % determine the height of the final table 
    cellHeightPowerBI.(string(variationType)) = 1;
    for currentSurvey = surveyList
        for headerGroupNum = 1:length(outputDataTablesCell.(string(tableTypeListMod(1))).(string(currentSurvey)))
            cellHeightPowerBI.(string(variationType)) = cellHeightPowerBI.(string(variationType)) + (width(outputDataTablesCell.(string(tableTypeListMod(1))).(string(currentSurvey)){headerGroupNum}) * (height(outputDataTablesCell.(string(tableTypeListMod(1))).(string(currentSurvey)){headerGroupNum}) - 1));
        end
    end
    % determine the width of the final table
    switch string(variationType)
        case 'Standard'
            powerBIPrefixHeaders.(string(variationType)) = {'Survey Name','Phase (QuestID)','Page Name'};
        case 'Weekday'
            powerBIPrefixHeaders.(string(variationType)) = {'Survey Name','Phase (QuestID)','Page Name','Weekday'};
        case 'PostalState'
            powerBIPrefixHeaders.(string(variationType)) = {'Survey Name','Phase (QuestID)','Page Name','PostalState'};
    end
    cellWidthPowerBI.(string(variationType)) = length(powerBIPrefixHeaders.(string(variationType))) + length(tableTypeList);
    
    % initialize the big table
    outputDataBigTableCellPowerBI.(string(variationType)) = cell(cellHeightPowerBI.(string(variationType)), cellWidthPowerBI.(string(variationType)));
    % outputDataBigTableCellPowerBI.(string(variationType))(1,:) = [{'Survey Name','Phase (QuestID)','Page Name','Weekday'}, tableTypeList]; 
    % outputDataBigTableCellPowerBI.(string(variationType))(1,:) = [{'Survey Name','Phase (QuestID)','Page Name'}, tableTypeList]; 
    outputDataBigTableCellPowerBI.(string(variationType))(1,:) = [powerBIPrefixHeaders.(string(variationType)), tableTypeListMod]; 
    for tableType = tableTypeListMod
        rowIndexPowerBI.(string(variationType)).(string(tableType)) = 2;
    end
    
    for currentSurvey = surveyList
        for tableType = tableTypeListMod
            for headerGroupNum = 1:length(outputDataTablesCell.(string(tableType)).(string(currentSurvey)))
                for phaseNum = 2:height(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum})
                    % insert survey name
                    outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Survey Name')) = deal(currentSurvey);
                    % insert phaseID
                    if strcmp(currentSurvey, 'Enrollment')
                        outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Phase (QuestID)')) = deal({sprintf('EnrollDup (%s)', string(extractBetween(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, 1), '/', whitespacePattern)))});
                        outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Phase (QuestID)')) = deal({sprintf('EnrollDup (%s)', string(extractBetween(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, 1), '/', whitespacePattern)))});
                    else
                        outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Phase (QuestID)')) = deal({sprintf('%s (%s)', string(extract(extractBefore(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, 1), '/'), digitsPattern)), string(extractBetween(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, 1), '/', whitespacePattern)))});
                        outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Phase (QuestID)')) = deal({sprintf('%s (%s)', string(extract(extractBefore(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, 1), '/'), digitsPattern)), string(extractBetween(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, 1), '/', whitespacePattern)))});
                    end
                    % insert 'Surveys Sent' page name before adding actual page names
                    outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)), matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')) = {'Surveys Sent'};
                    % insert actual page names
                    outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) + 1 : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')) = outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(1, 2:end)';
                    % insert data  
                    outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), tableType)) = outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, :)';
                    
                    % insert the day of the week or state abbr
                    if ~strcmp(variationType, 'Standard')
                        % insert the day of the week or state abbr
                        outputDataBigTableCellPowerBI.(string(variationType))(rowIndexPowerBI.(string(variationType)).(string(tableType)) : rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}) - 1, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), variationType)) = extractAfter(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum}(phaseNum, 1), '- ');
                    end
                    
                    rowIndexPowerBI.(string(variationType)).(string(tableType)) = rowIndexPowerBI.(string(variationType)).(string(tableType)) + width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)){headerGroupNum});
                end
            end
        end
    end

    % --- Format the tables further --- 
    % if tableType isn't Standard, remove variationType from end of column headers
    if ~strcmp(variationType, 'Standard')
        outputDataBigTableCellPowerBI.(string(variationType))(1, length(powerBIPrefixHeaders.(string(variationType))) + 1 : end) = tableTypeList;
    end
    % - move phaseID/questID (n=####) row entries to a "Total Surveys" column -
    outputDataBigTableCellPowerBI.(string(variationType))(1, width(outputDataBigTableCellPowerBI.(string(variationType)))+1) = {'Total Surveys'};
    % move headers from first tableType column to "Total Surveys"
    outputDataBigTableCellPowerBI.(string(variationType))([false; cellfun(@ischar, outputDataBigTableCellPowerBI.(string(variationType))(2:end, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), tableTypeList(1))))], end) = outputDataBigTableCellPowerBI.(string(variationType))([false; cellfun(@ischar, outputDataBigTableCellPowerBI.(string(variationType))(2:end, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), tableTypeList(1))))], matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), tableTypeList(1)));
    % replace cells that had survey name info with NaNs
    outputDataBigTableCellPowerBI.(string(variationType))([false; cellfun(@ischar, outputDataBigTableCellPowerBI.(string(variationType))(2:end, end))], find(matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), tableTypeList(1))):end-1) = {NaN}; 
    
    % - rename (ex. Dup88/20250101 (n=6659)) to (ex. (n=6659)) -
    outputDataBigTableCellPowerBI.(string(variationType))([false; cellfun(@ischar, outputDataBigTableCellPowerBI.(string(variationType))(2:end, end))], end) = extract(outputDataBigTableCellPowerBI.(string(variationType))([false; cellfun(@ischar, outputDataBigTableCellPowerBI.(string(variationType))(2:end, end))], end), '(n=' + digitsPattern + ')');
    
    % - replace empty cells in outputDataBigTableCellPowerBI with NaNs -
    outputDataBigTableCellPowerBI.(string(variationType))([false(1, width(outputDataBigTableCellPowerBI.(string(variationType)))); false(height(outputDataBigTableCellPowerBI.(string(variationType))) - 1, find(matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), tableTypeList(1))) - 1), cellfun(@isempty, outputDataBigTableCellPowerBI.(string(variationType))(2:end, find(matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), tableTypeList(1))):end))]) = {NaN};
    
    % - make tableTypeList headers in outputDataBigTableCellPowerBI pretty -
    outputDataBigTableCellPowerBI.(string(variationType))(1,contains(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'percent')) = join([outputDataBigTableCellPowerBI.(string(variationType))(1,contains(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'percent'))', repmat({'%'}, length(outputDataBigTableCellPowerBI.(string(variationType))(1,contains(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'percent'))), 1)], ' ')';
    outputDataBigTableCellPowerBI.(string(variationType))(1,:) = replace(outputDataBigTableCellPowerBI.(string(variationType))(1,:), {'percentPage','page'}, '');
    
    % - zero pad all single-digit pages in Page Name (i.e., [1] -> [01]) -
    outputDataBigTableCellPowerBI.(string(variationType))(startsWith(outputDataBigTableCellPowerBI.(string(variationType))(:, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')), '[' + digitsPattern + ']'), 3) = join([repmat({'['}, sum(startsWith(outputDataBigTableCellPowerBI.(string(variationType))(:, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')), '[' + digitsPattern + ']')), 1), ...
                                                                                                                                                                               strsplit(strtrim(sprintf('%02d ', str2double(string(extractBetween(outputDataBigTableCellPowerBI.(string(variationType))(startsWith(outputDataBigTableCellPowerBI.(string(variationType))(:, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')), '[' + digitsPattern + ']'), matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')), '[', ']'))))))', ...
                                                                                                                                                                               repmat({']'}, sum(startsWith(outputDataBigTableCellPowerBI.(string(variationType))(:, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')), '[' + digitsPattern + ']')), 1), ...
                                                                                                                                                                               extractAfter(outputDataBigTableCellPowerBI.(string(variationType))(startsWith(outputDataBigTableCellPowerBI.(string(variationType))(:, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')), '[' + digitsPattern + ']'), matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Page Name')), '[' + digitsPattern + ']')], '');
    
    
    % format all percentage values as a fracion of 1 (i.e., 69.0 -> 0.690)
    outputDataBigTableCellPowerBI.(string(variationType))(2:end, contains(outputDataBigTableCellPowerBI.(string(variationType))(1,:), '%')) = num2cell(cell2mat(outputDataBigTableCellPowerBI.(string(variationType))(2:end, contains(outputDataBigTableCellPowerBI.(string(variationType))(1,:), '%'))) / 100);
    
    % for Weekday, create a column at the end that represents the position
    % in the week for the corresponding weekday (e.g., Monday = 1, Saturday
    % = 6, etc.)
    if strcmp(variationType, 'Weekday')
        outputDataBigTableCellPowerBI.(string(variationType))(1, width(outputDataBigTableCellPowerBI.(string(variationType)))+1) = {'WeekdayIndex'};
        outputDataBigTableCellPowerBI.(string(variationType))(2:end, end) = replace(outputDataBigTableCellPowerBI.(string(variationType))(2:end, matches(outputDataBigTableCellPowerBI.(string(variationType))(1,:), 'Weekday')), weekdayList, string(1:length(weekdayList)));
    end
end
%%
% --- Calculate phase aggregates ---
% calculate aggregates for every combination of the following:
% Survey: [Enrollment, Long, Short]
% Period: [All Time (as a scalar), by Every Phase, by Every Week each Phase]
% Status: [Sent, Started, Completed]
statusList = {'Sent','Started','Completed'};
aggData = struct;
tic
for currentSurvey = surveyList
    for currentStatus = statusList
        % All Time (as a scalar)
        aggdata.(string(currentSurvey)).allTime.(string(currentStatus)) = 0; % initialize
        for phaseName = phaseList.(string(currentSurvey))
            if strcmp(currentStatus, 'Sent') % if looking at pageStarted or pageCompleted, use data corresponding to 'pg1'
                aggdata.(string(currentSurvey)).allTime.(string(currentStatus)) = aggdata.(string(currentSurvey)).allTime.(string(currentStatus)) + height(queryList.(string(queryListFieldNames(endsWith(queryListFieldNames, currentStatus)))).(string(currentSurvey)).data.(string(phaseName)));
            else
                aggdata.(string(currentSurvey)).allTime.(string(currentStatus)) = aggdata.(string(currentSurvey)).allTime.(string(currentStatus)) + height(queryList.(string(queryListFieldNames(endsWith(queryListFieldNames, currentStatus)))).(string(currentSurvey)).data.(string(phaseName)).pg1);
            end
        end

        % By Every Phase
        aggdata.(string(currentSurvey)).byPhase.(string(currentStatus)) = cell(length(phaseList.(string(currentSurvey))) + 1, 2); % initialize (first column is phaseName, second column is data)
        aggdata.(string(currentSurvey)).byPhase.(string(currentStatus))(1,:) = {'PhaseName','Data'};
        for phaseNum = 1:length(phaseList.(string(currentSurvey)))
            if strcmp(currentStatus, 'Sent') % if looking at pageStarted or pageCompleted, use data corresponding to 'pg1'
                aggdata.(string(currentSurvey)).byPhase.(string(currentStatus))(phaseNum + 1, :) = [phaseList.(string(currentSurvey))(phaseNum), height(queryList.(string(queryListFieldNames(endsWith(queryListFieldNames, currentStatus)))).(string(currentSurvey)).data.(string(phaseList.(string(currentSurvey))(phaseNum))))];
            else
                aggdata.(string(currentSurvey)).byPhase.(string(currentStatus))(phaseNum + 1, :) = [phaseList.(string(currentSurvey))(phaseNum), height(queryList.(string(queryListFieldNames(endsWith(queryListFieldNames, currentStatus)))).(string(currentSurvey)).data.(string(phaseList.(string(currentSurvey))(phaseNum))).pg1)];
            end
        end

        % By Every Week each Phase
        % determine how many weeks total there are for all the phases
        totalWeeks = sum(ceil(str2num(char(extract(cellstr(between(datetime(queryList.availablePhases.(string(currentSurvey)).data.StartDate, 'InputFormat','yyyy-MM-dd hh:mm:ss.SSS'), datetime(queryList.availablePhases.(string(currentSurvey)).data.EndDate, 'InputFormat','yyyy-MM-dd hh:mm:ss.SSS'), 'Days')), digitsPattern))) / 7));
        aggdata.(string(currentSurvey)).byWeek.(string(currentStatus)) = cell(totalWeeks + 1, 3); % initialize (first column is phaseName, second column is the date for the start of the week, third column is data)
        aggdata.(string(currentSurvey)).byWeek.(string(currentStatus))(1,:) = {'PhaseName','WeekStartDate','Data'};
        weekNum = 2;
        for phaseNum = 1:length(phaseList.(string(currentSurvey)))
            % generate the list of week starts and ends for each phase
            startDateList = datetime(queryList.availablePhases.(string(currentSurvey)).data.StartDate(phaseNum), 'InputFormat','yyyy-MM-dd hh:mm:ss.SSS', 'Format','yyyy-MM-dd') : calweeks(1) : datetime(queryList.availablePhases.(string(currentSurvey)).data.EndDate(phaseNum), 'InputFormat','yyyy-MM-dd hh:mm:ss.SSS', 'Format','yyyy-MM-dd');
            endDateList = [[[datetime(queryList.availablePhases.(string(currentSurvey)).data.StartDate(phaseNum), 'InputFormat','yyyy-MM-dd hh:mm:ss.SSS') + calweeks(1): calweeks(1) : datetime(queryList.availablePhases.(string(currentSurvey)).data.EndDate(phaseNum), 'InputFormat','yyyy-MM-dd hh:mm:ss.SSS')] - caldays(1)], datetime(queryList.availablePhases.(string(currentSurvey)).data.EndDate(phaseNum), 'InputFormat','yyyy-MM-dd hh:mm:ss.SSS')];
            for dateSetNum = 1:length(startDateList)
                if strcmp(currentStatus, 'Sent') % if looking at pageStarted or pageCompleted, use data corresponding to 'pg1'
                    aggdata.(string(currentSurvey)).byWeek.(string(currentStatus))(weekNum, :) = [phaseList.(string(currentSurvey))(phaseNum), char(startDateList(dateSetNum)), sum(isbetween(datetime(extract(queryList.(string(queryListFieldNames(endsWith(queryListFieldNames, currentStatus)))).(string(currentSurvey)).data.(string(phaseList.(string(currentSurvey))(phaseNum))).ActivityDate, digitsPattern(4) + '-' + digitsPattern(2) + '-' + digitsPattern(2)), 'InputFormat','yyyy-MM-dd'), startDateList(dateSetNum), endDateList(dateSetNum)))];
                else
                    aggdata.(string(currentSurvey)).byWeek.(string(currentStatus))(weekNum, :) = [phaseList.(string(currentSurvey))(phaseNum), char(startDateList(dateSetNum)), sum(isbetween(datetime(extract(queryList.(string(queryListFieldNames(endsWith(queryListFieldNames, currentStatus)))).(string(currentSurvey)).data.(string(phaseList.(string(currentSurvey))(phaseNum))).pg1.ActivityDate, digitsPattern(4) + '-' + digitsPattern(2) + '-' + digitsPattern(2)), 'InputFormat','yyyy-MM-dd'), startDateList(dateSetNum), endDateList(dateSetNum)))];
                end
                weekNum = weekNum + 1;
            end
        end
        fprintf('Aggregated %s data for %s survey...\n', string(currentStatus), string(currentSurvey));
    end
end
toc

%%%%%% FIGURE OUT WHY ActivityDate IS THE SAME FOR SENT, STARTED, AND COMPLETED %%%%%%

%%
% --- Export data to spreadsheet for Power BI ---
exportFileNamePowerBI = 'Dupuytrens KPI for Power BI.xlsx';
for variationType = variationTypeList
    writecell(outputDataBigTableCellPowerBI.(string(variationType)), exportFileNamePowerBI, 'Sheet', sprintf('KPI - %s', string(variationType)));
end


% -------------------------------------------------------------------------
function queryParts = evalQuery(inputStruct, title, description, query, dbConnection, databasePrefix1, databasePrefix2)
    queryParts = inputStruct;
    switch nargin
        case 5
            queryParts.title = title;
            queryParts.description = description;
            queryParts.query = query;
            % set which database to use
            execute(dbConnection, string(extract(queryParts.query, 'USE ' + lettersPattern)));
            % run the body of the query, with comments removed (i.e., remove
            % text between -- and --)
            queryParts.data = fetch(dbConnection, string(replace(extractAfter(queryParts.query, 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
        case 6
            queryParts.title = title;
            queryParts.description = description;
            queryParts.query.(string(databasePrefix1)) = query;
            % set which database to use
            execute(dbConnection, string(extract(queryParts.query.(string(databasePrefix1)), 'USE ' + lettersPattern)));
            % run the body of the query, with comments removed (i.e., remove
            % text between -- and --)
            queryParts.data.(string(databasePrefix1)) = fetch(dbConnection, string(replace(extractAfter(queryParts.query.(string(databasePrefix1)), 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
        case 7
            queryParts.title = title;
            queryParts.description = description;
            queryParts.query.(string(databasePrefix1)).(string(databasePrefix2)) = query;
            % set which database to use
            execute(dbConnection, string(extract(queryParts.query.(string(databasePrefix1)).(string(databasePrefix2)), 'USE ' + lettersPattern)));
            % run the body of the query, with comments removed (i.e., remove
            % text between -- and --)
            queryParts.data.(string(databasePrefix1)).(string(databasePrefix2)) = fetch(dbConnection, string(replace(extractAfter(queryParts.query.(string(databasePrefix1)).(string(databasePrefix2)), 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
        otherwise
            error('No case to handle %d input arguments', nargin);
    end
end

function rgbArray = colorMapCalc(percentVal, customCMap)
    % this function maps percentVal, a decimal between 0-1, onto a color map 
    % specified by customCMap, an nx3 array, where each row contains an rgb
    % triplet [0-255]. percentVal = 0 -> customCMap(1,:). percentVal = 1 
    % -> customCMap(n,:).
    
    if isnan(percentVal)
        rgbArray = [255, 255, 255];
    else
        % determine where on the color map percentVal falls
        cmapLoc = ((height(customCMap)-1) * percentVal) + 1;
        if round(cmapLoc) == cmapLoc
            rgbArray = customCMap(cmapLoc, :);
        else
            for rgbIndex = 1:3
                rgbArray(rgbIndex) = round(interp1([floor(cmapLoc), ceil(cmapLoc)], [customCMap(floor(cmapLoc), rgbIndex), customCMap(ceil(cmapLoc), rgbIndex)], cmapLoc));
            end
        end
    end
end

function range = excelRangeFinder(startRow,endRow,startCol,endCol)
    range = sprintf('%s%d:%s%d',colLetterConverter(startCol),startRow,colLetterConverter(endCol),endRow);

    function colLetter = colLetterConverter(colNum)
        colLetter = char(strrep(cellstr(char([floor((colNum-1)/689)',rem(floor((colNum-1)/26)'-1,26)+1,rem(colNum-1,26)'+1]+64)),'@',''));
    end
end

function colorVal = excelInteriorColorConverter(b)
    % b is an array of the form [b1, b2, b3], where b1, b2, and b3 are
    % within the range [0,255]
    colorVal = sum([256^0, 256^1, 256^2] .* double(b));
end






















