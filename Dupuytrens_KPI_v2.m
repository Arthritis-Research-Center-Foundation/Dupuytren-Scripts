% -------------------------------------------------------------------------
% This script is designed to query Dupuytrens data for Dr. Eaton, whose
% request is "Is there an automated way for me to see how many new 
% enrollees there have been during each phase? To develop my enrollment 
% and follow-up outreach efforts, I want to track the number of new 
% enrollees and phase participants in real time to track how well my 
% own outreach efforts are.". THIS SCRIPT ADDS THE CALCULATIONS OF THE 
% PERCENT OF PARTICIPANTS THAT COMPLETE EACH SURVEY PAGE. 02/03/2025
% -------------------------------------------------------------------------
% ------ Standard KIP Section ------

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
try
    queryIndex = 1;
    % --- Define queries ---
    % - Enrollment Questionnaires -
    % Enrollment: Total Started this Week - This is the total number of people who have started a Dupuytren Enrollment questionnaire this week
    queryList{queryIndex}.title = 'Enrollment: Total Started this Week';
    queryList{queryIndex}.description = 'This is the total number of people who have started a Dupuytren Enrollment questionnaire this week';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Name = ''Dupuytrens Enrollment'' -- must be enrolled in Dupuytrens --' ...
                                    '        AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --'];
    queryIndex = queryIndex + 1;

    % Enrollment: Total Completed this Week - Same as the above, but for a Completed status
    queryList{queryIndex}.title = 'Enrollment: Total Completed this Week';
    queryList{queryIndex}.description = 'This is the total number of people who have completed a Dupuytren Enrollment questionnaire this week';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Name = ''Dupuytrens Enrollment'' -- must be enrolled in Dupuytrens --' ...
                                    '        AND s.IsComplete = 1 -- must have a status of complete --' ...
                                    '        AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --'];
    queryIndex = queryIndex + 1;

    % Enrollment: Total Started Ever - We are looking for the total number of people who have a Dupuytren Diagnosis and have ever started a Dupuytren Enrollment
    queryList{queryIndex}.title = 'Enrollment: Total Started Ever';
    queryList{queryIndex}.description = 'The total number of people who have a Dupuytren Diagnosis and have ever started a Dupuytren Enrollment';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Name = ''Dupuytrens Enrollment'' -- must be enrolled in Dupuytrens --'];
    queryIndex = queryIndex + 1;

    % Enrollment: Total Completed Ever - Same as the above, but for a Completed status
    queryList{queryIndex}.title = 'Enrollment: Total Completed Ever';
    queryList{queryIndex}.description = 'The total number of people who have a Dupuytren Diagnosis and have ever completed a Dupuytren Enrollment';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Name = ''Dupuytrens Enrollment'' -- must be enrolled in Dupuytrens --' ...
                                    '        AND s.IsComplete = 1 -- must have a status of complete --'];    
    queryIndex = queryIndex + 1;

    % - Long Questionnaires -
    % Long Questionnaire: Current Phase Questionnaires Sent - This is the total number of phase questionnaires that have been sent out to Dupuytren participants this phase
    queryList{queryIndex}.title = 'Long Questionnaire: Current Phase Questionnaires Sent';
    queryList{queryIndex}.description = 'This is the total number of long phase questionnaires that have been sent out to Dupuytren participants this phase';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren)''-- specify the current phase name... --' ...
                                    '    AND up.StatusCode = ph.SentStatus -- confirm that the sent status is ''sent'' --'];    
    queryIndex = queryIndex + 1;

    % Long Questionnaire: Current Phase Questionnaires Started - This is the total number of phase questionnaires that have been started by Dupuytren participants this phase
    queryList{queryIndex}.title = 'Long Questionnaire: Current Phase Questionnaires Started';
    queryList{queryIndex}.description = 'This is the total number of long phase questionnaires that have been started by Dupuytren participants this phase';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren)''-- specify the current phase name... --' ...
		                            '    AND up.StatusCode = ph.StartedStatus -- confirm that the started status is ''started'' --'];    
    queryIndex = queryIndex + 1;

    % Long Questionnaire: Current Phase Questionnaires Completed - This is the total number of phase questionnaires that have been completed by Dupuytren participants this phase
    queryList{queryIndex}.title = 'Long Questionnaire: Current Phase Questionnaires Completed';
    queryList{queryIndex}.description = 'This is the total number of long phase questionnaires that have been completed by Dupuytren participants this phase';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren)''-- specify the current phase name... --' ...
		                            '    AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is ''complete'' --'];    
    queryIndex = queryIndex + 1;

    % Long Questionnaire: Phase Questionnaires Completed this Week - This is the total number of phase questionnaires that have been completed by Dupuytren participants this week
    queryList{queryIndex}.title = 'Long Questionnaire: Phase Questionnaires Completed this Week';
    queryList{queryIndex}.description = 'This is the total number of long phase questionnaires that have been completed by Dupuytren participants this week';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren)''-- specify the current phase name... --' ...
		                            '    AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is ''complete'' --' ...
                                    '    AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --'];    
    queryIndex = queryIndex + 1;

    % Long Questionnaire: Total Phase Questionnaires Ever Sent - This is the total number of phase questionnaires that have ever been sent out to Dupuytren participants
    queryList{queryIndex}.title = 'Long Questionnaire: Total Phase Questionnaires Ever Sent';
    queryList{queryIndex}.description = 'This is the total number of long phase questionnaires that have ever been sent out to Dupuytren participants';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Name Like ''%Questionnaire (%Dupuytren%'' -- specify the long questionnaire phase name... --' ...
			                        '        AND ph.Name Not Like ''%Questionnaire (%Dupuytren%Short%'' ' ...
		                            '    AND up.StatusCode = ph.SentStatus -- confirm that the sent status is ''sent'' --'];    
    queryIndex = queryIndex + 1;

    % Long Questionnaire: Total Phase Questionnaires Ever Completed - This is the total number of phase questionnaires that have ever been completed by Dupuytren participants
    queryList{queryIndex}.title = 'Long Questionnaire: Total Phase Questionnaires Ever Completed';
    queryList{queryIndex}.description = 'This is the total number of long phase questionnaires that have ever been completed by Dupuytren participants';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Name Like ''%Questionnaire (%Dupuytren%'' -- specify the long questionnaire phase name... --' ...
			                        '        AND ph.Name Not Like ''%Questionnaire (%Dupuytren%Short%'' ' ...
		                            '    AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is ''complete'' --'];    
    queryIndex = queryIndex + 1;

    % - Short Questionnaires -
    % Short Questionnaire: Current Phase Questionnaires Sent - This is the total number of phase questionnaires that have been sent out to Dupuytren participants this phase
    queryList{queryIndex}.title = 'Short Questionnaire: Current Phase Questionnaires Sent';
    queryList{queryIndex}.description = 'This is the total number of short phase questionnaires that have been sent out to Dupuytren participants this phase';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren Short)''-- specify the current phase name... --' ...
		                            '    AND up.StatusCode = ph.SentStatus -- confirm that the sent status is ''sent'' --'];    
    queryIndex = queryIndex + 1;

    % Short Questionnaire: Current Phase Questionnaires Started - This is the total number of phase questionnaires that have been started by Dupuytren participants this phase
    queryList{queryIndex}.title = 'Short Questionnaire: Current Phase Questionnaires Started';
    queryList{queryIndex}.description = 'This is the total number of short phase questionnaires that have been started by Dupuytren participants this phase';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren Short)''-- specify the current phase name... --' ...
		                            '    AND up.StatusCode = ph.StartedStatus -- confirm that the started status is ''started'' --'];    
    queryIndex = queryIndex + 1;

    % Short Questionnaire: Current Phase Questionnaires Completed - This is the total number of phase questionnaires that have been completed by Dupuytren participants this phase
    queryList{queryIndex}.title = 'Short Questionnaire: Current Phase Questionnaires Completed';
    queryList{queryIndex}.description = 'This is the total number of short phase questionnaires that have been completed by Dupuytren participants this phase';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren Short)''-- specify the current phase name... --' ...
		                            '    AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is ''complete'' --'];    
    queryIndex = queryIndex + 1;

    % Short Questionnaire: Phase Questionnaires Completed this Week - This is the total number of phase questionnaires that have been completed by Dupuytren participants this week
    queryList{queryIndex}.title = 'Short Questionnaire: Phase Questionnaires Completed this Week';
    queryList{queryIndex}.description = 'This is the total number of short phase questionnaires that have been completed by Dupuytren participants this week';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
                                    '        AND ph.Active = 1 -- specify active surveys --' ...
                                    '        AND ph.Name Like ''%Questionnaire (Dupuytren Short)''-- specify the current phase name... --' ...
		                            '    AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is ''complete'' --' ...
                                    '    AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --'];    
    queryIndex = queryIndex + 1;

    % Short Questionnaire: Total Phase Questionnaires Ever Sent - This is the total number of phase questionnaires that have ever been sent out to Dupuytren participants
    queryList{queryIndex}.title = 'Short Questionnaire: Total Phase Questionnaires Ever Sent';
    queryList{queryIndex}.description = 'This is the total number of short phase questionnaires that have ever been sent out to Dupuytren participants';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
			                        '        AND ph.Name Like ''%Questionnaire (Dupuytren Short)'' -- specify the short questionnaire phase name... --' ...
		                            '    AND up.StatusCode = ph.SentStatus -- confirm that the sent status is ''sent'' --'];    
    queryIndex = queryIndex + 1;

    % Short Questionnaire: Total Phase Questionnaires Ever Completed - This is the total number of phase questionnaires that have ever been completed by Dupuytren participants
    queryList{queryIndex}.title = 'Short Questionnaire: Total Phase Questionnaires Ever Completed';
    queryList{queryIndex}.description = 'This is the total number of short phase questionnaires that have ever been completed by Dupuytren participants';
    queryList{queryIndex}.query = [ 'USE Forward ' ...
                                    'SELECT COUNT(*) as Count ' ...
                                    '    FROM arc.dbo.Natalog n ' ...
                                    '    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
                                    '    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
                                    '    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
                                    '    WHERE diagnosis = ''DUP'' -- must have DUP diagnosis --' ...
			                        '        AND ph.Name Like ''%Questionnaire (Dupuytren Short)'' -- specify the short questionnaire phase name... --' ...
		                            '    AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is ''complete'' --'];    
    

    % for each query...
    for queryNum = 1:length(queryList)
        % set which database to use
        if (queryNum == 1) || ~strcmp(string(extract(queryList{queryNum}.query, 'USE ' + lettersPattern)), string(extract(queryList{queryNum-1}.query, 'USE ' + lettersPattern)))
            execute(dbConnection, string(extract(queryList{queryNum}.query, 'USE ' + lettersPattern)));
        end
        % run the body of the query, with comments removed (i.e., remove
        % text between -- and --)
        queryList{queryNum}.data = fetch(dbConnection, string(replace(extractAfter(queryList{queryNum}.query, 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
    end
    close(dbConnection);
catch ME
    close(dbConnection);
    rethrow(ME);
end

%%
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
writecell(outputData, 'Dupuytrens KPI.xlsx', 'Sheet','phaseAggregates')

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

% --- Define default queries ---
defaultQueryList.availableTables.description = 'The list of unique survey phases in Dupuytrens';
defaultQueryList.availableTables.query = [   'USE Dupuytrens ' ...
                                             'SELECT TABLE_NAME ' ...
                                             '    FROM Dupuytrens.INFORMATION_SCHEMA.TABLES ' ...
                                             '    WHERE TABLE_NAME like ''{SURVEYNAME}%_pg%'' -- must have {SURVEYNAME} in table name --']; % SURVEYNAME = {'Enroll', 'Dup', 'Short'}


defaultQueryList.pageHeaders.description = 'The list of page numbers and page titles ALONG WITH the first survey that started using that specific order of page titles and numbers';
defaultQueryList.pageHeaders.query = [   'USE Forward ' ...
                                         'SELECT distinct DatabasePrefix, ''['' + CAST(PageNumber as VARCHAR(3)) + ''] '' + Pageheader as ''Page Header'' ' ...
                                         '    FROM Page p ' ...
                                         '    JOIN SurveyTool st ON st.SurveyToolId = p.SurveyToolId -- match SurveyToolIDs between Forward.SurveyTools and Forward.Page --' ...
                                         '    JOIN Phase ph ON ph.PhaseId = st.PhaseId -- match phase IDs between Forward.Phase and Forward.SurveyTools --' ...
                                         '    WHERE ph.DatabasePrefix like ''{SURVEYNAME}%'' -- must have {SURVEYNAME} database prefix --' ...
		                                 '        AND ProjectId = 46 -- specify the projectID... --']; % SURVEYNAME = {'Enroll', 'Dup', 'Short'}

defaultQueryList.questionnairesSent.description = 'The number of people who should have completed questionnaires (i.e., the number of questionnaires sent out)';
defaultQueryList.questionnairesSent.query = [  'USE Forward ' ...
                                               'SELECT COUNT(*) AS ''Questionnaires sent out'' ' ...
	                                           'FROM UserPhase up -- grab user activity --' ...
	                                           'JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
	                                           'JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
	                                           'JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
	                                           'WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
		                                       '    AND pj.Name = ''Dupuytrens'' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...
		                                       '    AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
                                                                                                                                        % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}

defaultQueryList.pageCompleted.description = 'The number of people who completed the page';
defaultQueryList.pageCompleted.query = [ 'USE Forward ' ...
                                         'SELECT COUNT(*) AS ''Num Participants that Started Page'' ' ...
	                                     'FROM Dupuytrens.dbo.{SURVEYPAGEHEADER} q -- grab info for specific page of specific questionnaire --' ...
                                         'JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
	                                     'JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
	                                     'JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
	                                     'JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
	                                     'WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
		                                 '    AND pj.Name = ''Dupuytrens'' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...
		                                 '    AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --', ...
		                                 '    AND Complete = 1 -- only include complete pages --']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
                                                                                                    % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
                                                                                                    % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}



%%
tic
try
    % --- Determine the list of page numbers and page titles ALONG WITH the
    % first survey that started using that specific order of page titles
    % and numbers for Enrollment, Long Surveys, and Short Surveys ---
    % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgMeds'))
    % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
    % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}
    for currentSurvey = {'Enrollment', 'Long', 'Short'}
        switch string(currentSurvey)
            case 'Enrollment'
                surveyName = 'Enroll';
                surveySelect = 'ph.Name LIKE ''%Enrollment%''';
                surveyDescription = 'Dupuytrens Enrollment PhaseId';
                surveyPageHeader = [sprintf('%s_pg', 'EnrollDup'), '%d'];
            case 'Long'
                surveyName = 'Dup';
                surveySelect = 'ph.DatabasePrefix = ''Dup%d''';
                surveyDescription = 'Dupuytrens Long Survey for ph%d';
                surveyPageHeader = [sprintf('%s_pg', 'Dup%d'), '%d'];
            case 'Short'
                surveyName = 'Short';
                surveySelect = 'ph.DatabasePrefix = ''Short%d''';
                surveyDescription = 'Dupuytrens Short Survey for ph%d';
                surveyPageHeader = [sprintf('%s_pg', 'Short%d'), '%d'];
        end
        queryList.pageHeaders.(string(currentSurvey)).title = sprintf('%s Questionnaire: PageHeaders', string(currentSurvey));
        queryList.pageHeaders.(string(currentSurvey)).description = defaultQueryList.pageHeaders.description;
        queryList.pageHeaders.(string(currentSurvey)).query = strrep(defaultQueryList.pageHeaders.query, '{SURVEYNAME}', surveyName);
        
        % - Execute query -
        % set which database to use
        execute(dbConnection, string(extract(queryList.pageHeaders.(string(currentSurvey)).query, 'USE ' + lettersPattern)));

        % run the body of the query, with comments removed (i.e., remove
        % text between -- and --)
        queryList.pageHeaders.(string(currentSurvey)).data = fetch(dbConnection, string(replace(extractAfter(queryList.pageHeaders.(string(currentSurvey)).query, 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));

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
        for databasePrefixName = unique(queryList.pageHeaders.(string(currentSurvey)).data.DatabasePrefix)'
            databasePrefixNumList.(string(currentSurvey)) = [databasePrefixNumList.(string(currentSurvey)), str2double(extract(databasePrefixName, digitsPattern))];
            pageHeaders.(string(currentSurvey)).(string(databasePrefixName)) = queryList.pageHeaders.(string(currentSurvey)).data.PageHeader(matches(queryList.pageHeaders.(string(currentSurvey)).data.DatabasePrefix, databasePrefixName));
        end
        
        if ~strcmp(currentSurvey, 'Enrollment')
            % - Determine the first and last available phase for the current survey -
            queryList.availableTables.(string(currentSurvey)).title = sprintf('%s Questionnaire: Available Tables', string(currentSurvey));
            queryList.availableTables.(string(currentSurvey)).description = defaultQueryList.availableTables.description;
            queryList.availableTables.(string(currentSurvey)).query = strrep(defaultQueryList.availableTables.query, '{SURVEYNAME}', surveyName);
            execute(dbConnection, string(extract(queryList.availableTables.(string(currentSurvey)).query, 'USE ' + lettersPattern)));
            queryList.availableTables.(string(currentSurvey)).data = fetch(dbConnection, string(replace(extractAfter(queryList.availableTables.(string(currentSurvey)).query, 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
            
            firstPhaseNum.(string(currentSurvey)) = min(unique(str2double(extract(extractBefore(queryList.availableTables.(string(currentSurvey)).data.TABLE_NAME, '_'), digitsPattern))));
            lastPhaseNum.(string(currentSurvey)) = max(unique(str2double(extract(extractBefore(queryList.availableTables.(string(currentSurvey)).data.TABLE_NAME, '_'), digitsPattern))));

            % change the first database prefix to firstSurveyNum
            pageHeaders.(string(currentSurvey)) = cell2struct(struct2cell(pageHeaders.(string(currentSurvey))), cellstr(strrep(fieldnames(pageHeaders.(string(currentSurvey))),  string(databasePrefixNumList.(string(currentSurvey))(1)), string(firstPhaseNum.(string(currentSurvey)))))');
            databasePrefixNumList.(string(currentSurvey))(1) = firstPhaseNum.(string(currentSurvey));

            % starting at firstSurveyNum
            for phaseNum = firstPhaseNum.(string(currentSurvey)):lastPhaseNum.(string(currentSurvey))
                % - Total questionnaires sent -
                queryList.questionnairesSent.(string(currentSurvey)).title = sprintf('%s Questionnaire: Total Sent', string(currentSurvey));
                queryList.questionnairesSent.(string(currentSurvey)).description = defaultQueryList.questionnairesSent.description;
                queryList.questionnairesSent.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)) = sprintf(strrep(strrep(defaultQueryList.questionnairesSent.query, '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), phaseNum, phaseNum);
                execute(dbConnection, string(extract(queryList.questionnairesSent.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)), 'USE ' + lettersPattern)));
                queryList.questionnairesSent.(string(currentSurvey)).data.(sprintf('%s%d', surveyName,phaseNum)) = fetch(dbConnection, string(replace(extractAfter(queryList.questionnairesSent.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)), 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
                % determine which database prefix's page headers to use
                headerSetPrefixNum = databasePrefixNumList.(string(currentSurvey))(phaseNum >= databasePrefixNumList.(string(currentSurvey)) & [phaseNum < databasePrefixNumList.(string(currentSurvey))(2:end), true]);
                for pageNum = str2double(extractBetween(pageHeaders.(string(currentSurvey)).(sprintf('%s%d', surveyName, headerSetPrefixNum)), '[', ']'))'
                    currentSurveyDescription = sprintf(surveyDescription, pageNum);
                    currentSurveyPageHeader = sprintf(surveyPageHeader, phaseNum, pageNum);
                    % - Total patients who completed the current page -
                    queryList.pageCompleted.(string(currentSurvey)).title = sprintf('%s Questionnaire: Total Completed', string(currentSurvey));
                    queryList.pageCompleted.(string(currentSurvey)).description = defaultQueryList.pageCompleted.description;
                    if ~any(matches(queryList.availableTables.(string(currentSurvey)).data.TABLE_NAME, sprintf('%s%d_pg%d', surveyName, phaseNum, pageNum))) % if current table name can't be found in Dupuytrens tables, then the current page must be pgDrug
                        % queryList.pageCompleted.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)).(sprintf('pg%d', pageNum)) = sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageStarted.query, '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), 'pg%d', 'pgDrug'), phaseNum, phaseNum, phaseNum, phaseNum);
                        queryList.pageCompleted.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)).(sprintf('pg%d', pageNum)) = sprintf(strrep(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), 'pg%d', 'pgDrug'), phaseNum, phaseNum, phaseNum, phaseNum);
                    else
                        % queryList.pageCompleted.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)).(sprintf('pg%d', pageNum)) = sprintf(strrep(strrep(strrep(defaultQueryList.pageStarted.query, '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), phaseNum, pageNum, phaseNum, pageNum, phaseNum, phaseNum);
                        queryList.pageCompleted.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)).(sprintf('pg%d', pageNum)) = sprintf(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), phaseNum, pageNum, phaseNum, pageNum, phaseNum, phaseNum);
                    end
                    execute(dbConnection, string(extract(queryList.pageCompleted.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)).(sprintf('pg%d', pageNum)), 'USE ' + lettersPattern)));
                    queryList.pageCompleted.(string(currentSurvey)).data.(sprintf('%s%d', surveyName,phaseNum)).(sprintf('pg%d', pageNum)) = fetch(dbConnection, string(replace(extractAfter(queryList.pageCompleted.(string(currentSurvey)).query.(sprintf('%s%d', surveyName,phaseNum)).(sprintf('pg%d', pageNum)), 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
                    fprintf('%s Survey, Phase %d, Page %d completed...\n', string(currentSurvey), phaseNum, pageNum);
                end
            end
        else
            % - Determine the first and last available phase for the current survey -
            queryList.availableTables.(string(currentSurvey)).title = sprintf('%s Questionnaire: Available Tables', string(currentSurvey));
            queryList.availableTables.(string(currentSurvey)).description = defaultQueryList.availableTables.description;
            queryList.availableTables.(string(currentSurvey)).query = strrep(defaultQueryList.availableTables.query, '{SURVEYNAME}', surveyName);
            execute(dbConnection, string(extract(queryList.availableTables.(string(currentSurvey)).query, 'USE ' + lettersPattern)));
            queryList.availableTables.(string(currentSurvey)).data = fetch(dbConnection, string(replace(extractAfter(queryList.availableTables.(string(currentSurvey)).query, 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
            
            % - Total questionnaires sent -
            queryList.questionnairesSent.(string(currentSurvey)).title = sprintf('%s Questionnaire: Total Sent', string(currentSurvey));
            queryList.questionnairesSent.(string(currentSurvey)).description = defaultQueryList.questionnairesSent.description;
            queryList.questionnairesSent.(string(currentSurvey)).query.(extractBefore(surveyPageHeader, '_')) = strrep(strrep(defaultQueryList.questionnairesSent.query, '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription);
            execute(dbConnection, string(extract(queryList.questionnairesSent.(string(currentSurvey)).query.(extractBefore(surveyPageHeader, '_')), 'USE ' + lettersPattern)));
            queryList.questionnairesSent.(string(currentSurvey)).data.(extractBefore(surveyPageHeader, '_')) = fetch(dbConnection, string(replace(extractAfter(queryList.questionnairesSent.(string(currentSurvey)).query.(extractBefore(surveyPageHeader, '_')), 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
            
            for pageNum = str2double(extractBetween(pageHeaders.(string(currentSurvey)).(extractBefore(surveyPageHeader, '_')), '[', ']'))'
                % - Total patients who completed the current page -
                queryList.pageCompleted.(string(currentSurvey)).title = sprintf('%s Questionnaire: Total Completed', string(currentSurvey));
                queryList.pageCompleted.(string(currentSurvey)).description = defaultQueryList.pageCompleted.description;
                if ~any(matches(queryList.availableTables.(string(currentSurvey)).data.TABLE_NAME, sprintf('%s_pg%d', extractBefore(surveyPageHeader, '_'), pageNum))) % if current table name can't be found in Dupuytrens tables, then the current page must be pgDrug
                    queryList.pageCompleted.(string(currentSurvey)).query.(extractBefore(surveyPageHeader, '_')).(sprintf('pg%d', pageNum)) = strrep(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', surveySelect), '{SURVEYDESCRIPTION}', surveyDescription), 'pg%d', 'pgDrug');
                else
                    queryList.pageCompleted.(string(currentSurvey)).query.(extractBefore(surveyPageHeader, '_')).(sprintf('pg%d', pageNum)) = sprintf(strrep(strrep(strrep(defaultQueryList.pageCompleted.query, '{SURVEYPAGEHEADER}', surveyPageHeader), '{SURVEYSELECT}', strrep(surveySelect, '%', '%%')), '{SURVEYDESCRIPTION}', surveyDescription), pageNum, pageNum);
                end
                execute(dbConnection, string(extract(queryList.pageCompleted.(string(currentSurvey)).query.(extractBefore(surveyPageHeader, '_')).(sprintf('pg%d', pageNum)), 'USE ' + lettersPattern)));
                queryList.pageCompleted.(string(currentSurvey)).data.(extractBefore(surveyPageHeader, '_')).(sprintf('pg%d', pageNum)) = fetch(dbConnection, string(replace(extractAfter(queryList.pageCompleted.(string(currentSurvey)).query.(extractBefore(surveyPageHeader, '_')).(sprintf('pg%d', pageNum)), 'USE ' + lettersPattern + ' '), '--' + wildcardPattern + '--', '')));
                fprintf('%s Survey, Page %d completed...\n', string(currentSurvey), pageNum);
            end
        end

    end

catch ME
    close(dbConnection);
    rethrow(ME);
end
toc
%%
% --- Arrange the data into a neat table ---
for currentSurvey = {'Enrollment', 'Long', 'Short'}
    if ~strcmp(currentSurvey, 'Enrollment')
        % each pageHeader will have its own table (cus each pageHeader has
        % different pages). determine which phases go with which
        % pageHeaders
        % determine which phases go with current pageHeader
        phaseList.(string(currentSurvey)) = firstPhaseNum.(string(currentSurvey)):lastPhaseNum.(string(currentSurvey));
        pageHeaderNumList.(string(currentSurvey)) = str2double(extract(fieldnames(pageHeaders.(string(currentSurvey))), digitsPattern))';
        for pageHeaderNumIndex = 1:length(fieldnames(pageHeaders.(string(currentSurvey))))
            % determine height of table (num phases to include for
            % current pageHeader plus 1)
            if pageHeaderNumIndex < length(fieldnames(pageHeaders.(string(currentSurvey)))) 
                phasesToInclude = pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex) : pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex + 1) - 1; % phases between current pageHeader and next pageHeader
            else % if on last pageHeader...
                phasesToInclude = pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex) : lastPhaseNum.(string(currentSurvey)); % phases between the last pageHeader and the last phase (including the last phase)
            end
            % - Page Completed -
            % initialize pageCompleted table for current pageHeader
            outputDataTablesCell.pageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))) = cell(length(phasesToInclude) + 1, height(pageHeaders.(string(currentSurvey)).(sprintf('%s%d', string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))) + 1);
            outputDataTablesCell.pageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))(1, 2:end) = pageHeaders.(string(currentSurvey)).(sprintf('%s%d', string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))';
            outputDataTablesCell.pageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))(2:end, 1) = flipud(cellstr(join([repmat(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)), length(phasesToInclude), 1), string(phasesToInclude')], '')));
            % insert data for each page for the current pageHeader
            currentPhaseNameList = outputDataTablesCell.pageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))(2:end, 1)';
            pageNumList = str2double(extract(fieldnames(queryList.pageCompleted.(string(currentSurvey)).data.(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))), digitsPattern)');
            for currentPhaseNameIndex = 1:length(currentPhaseNameList)
                for pageNumIndex = 1:length(pageNumList)
                    outputDataTablesCell.pageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))(currentPhaseNameIndex + 1, pageNumIndex + 1) = table2cell(queryList.pageCompleted.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))));
                end
                % append number of questionnaires sent as (n=####) to phase 
                % name first column
                outputDataTablesCell.pageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))){currentPhaseNameIndex + 1, 1} = sprintf('%s (n=%d)', outputDataTablesCell.pageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))){currentPhaseNameIndex + 1, 1}, cell2mat(table2cell(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))))));
            end

            % - Percent Page Completed -
            % initialize percentPageCompleted table for current pageHeader
            outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))) = cell(length(phasesToInclude) + 1, height(pageHeaders.(string(currentSurvey)).(sprintf('%s%d', string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))) + 1);
            outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))(1, 2:end) = pageHeaders.(string(currentSurvey)).(sprintf('%s%d', string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))';
            outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))(2:end, 1) = flipud(cellstr(join([repmat(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)), length(phasesToInclude), 1), string(phasesToInclude')], '')));
            % insert data for each page for the current pageHeader
            currentPhaseNameList = outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))(2:end, 1)';
            pageNumList = str2double(extract(fieldnames(queryList.pageCompleted.(string(currentSurvey)).data.(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex)))), digitsPattern)');
            for currentPhaseNameIndex = 1:length(currentPhaseNameList)
                for pageNumIndex = 1:length(pageNumList)
                    outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))){currentPhaseNameIndex + 1, pageNumIndex + 1} = round(10000 * (table2array(queryList.pageCompleted.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex)))) / table2array(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex)))))) /100;
                end
                % append number of questionnaires sent as (n=####) to phase 
                % name first column
                outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))){currentPhaseNameIndex + 1, 1} = sprintf('%s (n=%d)', outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))){currentPhaseNameIndex + 1, 1}, cell2mat(table2cell(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))))));
            end
        end

    else
        % - Page Completed -
        % initialize pageCompleted table for current pageHeader
        outputDataTablesCell.pageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)))) = cell(2, height(pageHeaders.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))) + 1);
        outputDataTablesCell.pageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))(1, 2:end) = pageHeaders.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))';
        outputDataTablesCell.pageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))(2:end, 1) = unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern));
        % insert data for each page for the current pageHeader
        currentPhaseNameList = outputDataTablesCell.pageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))(2:end, 1)';
        pageNumList = str2double(extract(fieldnames(queryList.pageCompleted.(string(currentSurvey)).data.(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))), digitsPattern)');
        for currentPhaseNameIndex = 1:length(currentPhaseNameList)
            for pageNumIndex = 1:length(pageNumList)
                outputDataTablesCell.pageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))(currentPhaseNameIndex + 1, pageNumIndex + 1) = table2cell(queryList.pageCompleted.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex))));
            end
            % append number of questionnaires sent as (n=####) to phase 
            % name first column
            outputDataTablesCell.pageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)))){currentPhaseNameIndex + 1, 1} = sprintf('%s (n=%d)', outputDataTablesCell.pageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)))){currentPhaseNameIndex + 1, 1}, cell2mat(table2cell(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))))));
        end

        % - Percent Page Completed -
        % initialize percentPageCompleted table for current pageHeader
        outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)))) = cell(2, height(pageHeaders.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))) + 1);
        outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))(1, 2:end) = pageHeaders.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))';
        outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))(2:end, 1) = unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern));
        % insert data for each page for the current pageHeader
        currentPhaseNameList = outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))(2:end, 1)';
        pageNumList = str2double(extract(fieldnames(queryList.pageCompleted.(string(currentSurvey)).data.(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))))), digitsPattern)');
        for currentPhaseNameIndex = 1:length(currentPhaseNameList)
            for pageNumIndex = 1:length(pageNumList)
                outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)))){currentPhaseNameIndex + 1, pageNumIndex + 1} = round(10000 * (table2array(queryList.pageCompleted.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))).(sprintf('pg%d', pageNumList(pageNumIndex)))) / table2array(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex)))))) / 100;
            end
            % append number of questionnaires sent as (n=####) to phase 
            % name first column
            outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)))){currentPhaseNameIndex + 1, 1} = sprintf('%s (n=%d)', outputDataTablesCell.percentPageCompleted.(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern)))){currentPhaseNameIndex + 1, 1}, cell2mat(table2cell(queryList.questionnairesSent.(string(currentSurvey)).data.(string(currentPhaseNameList(currentPhaseNameIndex))))));
        end
    end
end

%% 
% --- Arrange the data tables into a singular table for placement into a
% spreadsheet ---
for currentSurvey = {'Enrollment', 'Long', 'Short'}
    for tableType = {'percentPageCompleted', 'pageCompleted'}
        if ~strcmp(currentSurvey, 'Enrollment')
            % determine the height of the big cell array "table"
            cellHeight = lastPhaseNum.(string(currentSurvey)) - firstPhaseNum.(string(currentSurvey)) + 1 + 2*(length(fieldnames(outputDataTablesCell.(string(tableType)).(string(currentSurvey))))) - 1;
            % determine the width of the big cell array "table"
            cellWidth = 0;
            for pageHeaderNumIndex = 1:length(fieldnames(pageHeaders.(string(currentSurvey))))
                cellWidth = max([cellWidth, width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(pageHeaderNumIndex))))]);
            end
            % initialize the big cell array "table"
            outputDataBigTableCell.(string(tableType)).(string(currentSurvey)) = cell(cellHeight, cellWidth);
            % insert individual tables from outputDataTablesCell into
            % outputDataBigTableCell
            for pageHeaderNumIndex = 1:length(fieldnames(pageHeaders.(string(currentSurvey))))
                % determine the top row that the current subtable should be
                % inserted into
                if pageHeaderNumIndex == 1
                    rowIndex.(string(tableType)).(string(currentSurvey))(1) = 1;
                else
                    rowIndex.(string(tableType)).(string(currentSurvey))(pageHeaderNumIndex) = rowIndex.(string(tableType)).(string(currentSurvey))(pageHeaderNumIndex - 1) + height(outputDataTablesCell.(string(tableType)).(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(length(fieldnames(pageHeaders.(string(currentSurvey)))) - pageHeaderNumIndex + 2)))) + 1;
                end
                % insert the subtable
                outputDataBigTableCell.(string(tableType)).(string(currentSurvey))(rowIndex.(string(tableType)).(string(currentSurvey))(pageHeaderNumIndex):rowIndex.(string(tableType)).(string(currentSurvey))(pageHeaderNumIndex) + height(outputDataTablesCell.(string(tableType)).(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(length(fieldnames(pageHeaders.(string(currentSurvey)))) - pageHeaderNumIndex + 1)))) - 1, 1:width(outputDataTablesCell.(string(tableType)).(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(length(fieldnames(pageHeaders.(string(currentSurvey)))) - pageHeaderNumIndex + 1))))) = outputDataTablesCell.(string(tableType)).(string(currentSurvey)).(sprintf('%s%d',string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))), pageHeaderNumList.(string(currentSurvey))(length(fieldnames(pageHeaders.(string(currentSurvey)))) - pageHeaderNumIndex + 1)));
            end

            % - Specify spreadsheet formatting -
            format.(string(tableType)).(string(currentSurvey)).boldRows = rowIndex.(string(tableType)).(string(currentSurvey));
            format.(string(tableType)).(string(currentSurvey)).bottomBorderRows = rowIndex.(string(tableType)).(string(currentSurvey));
            format.(string(tableType)).(string(currentSurvey)).rightBorderRows = 1:height(outputDataBigTableCell.(string(tableType)).(string(currentSurvey)));
            format.(string(tableType)).(string(currentSurvey)).rightBorderRows(ismember(format.(string(tableType)).(string(currentSurvey)).rightBorderRows, [rowIndex.(string(tableType)).(string(currentSurvey)), rowIndex.(string(tableType)).(string(currentSurvey))-1])) = [];
            if strcmp(tableType, 'percentPageCompleted')
                % calculate the cell fill colors for cells with numeric
                % data
                format.(string(tableType)).(string(currentSurvey)).isNumeric = cellfun(@isnumeric, outputDataBigTableCell.(string(tableType)).(string(currentSurvey))) & ~cellfun(@isempty, outputDataBigTableCell.(string(tableType)).(string(currentSurvey)));
                [format.(string(tableType)).(string(currentSurvey)).fillColorAbsolute, format.(string(tableType)).(string(currentSurvey)).fillColorRelative] = deal(NaN(size(outputDataBigTableCell.(string(tableType)).(string(currentSurvey)))));
                for rowNum = 1:height(format.(string(tableType)).(string(currentSurvey)).isNumeric)
                    for colNum = 2:width(format.(string(tableType)).(string(currentSurvey)).isNumeric)
                        if format.(string(tableType)).(string(currentSurvey)).isNumeric(rowNum, colNum)
                            % fill color where range is from 0-1
                            format.(string(tableType)).(string(currentSurvey)).fillColorAbsolute(rowNum, colNum) = excelInteriorColorConverter(colorMapCalc(outputDataBigTableCell.(string(tableType)).(string(currentSurvey)){rowNum, colNum} / 100, [255, 0, 0; 255, 255, 0; 0, 255, 0])); % cmap is [red, yellow, green]
                            % fill color where range is min(currentPhaseValues) - max(currentPhaseValues) 
                            % (i.e., map [0,1] onto [min(currentPhaseValues),max(currentPhaseValues)]
                            format.(string(tableType)).(string(currentSurvey)).fillColorRelative(rowNum, colNum) = excelInteriorColorConverter(colorMapCalc(interp1([min(cell2mat(outputDataBigTableCell.(string(tableType)).(string(currentSurvey))(rowNum, 2:end)) / 100), max(cell2mat(outputDataBigTableCell.(string(tableType)).(string(currentSurvey))(rowNum, 2:end)) / 100)], [0,1], outputDataBigTableCell.(string(tableType)).(string(currentSurvey)){rowNum, colNum} / 100), [255, 0, 0; 255, 255, 0; 0, 255, 0])); % cmap is [red, yellow, green]
                        end
                    end
                end
            end

        else
            % insert individual table from outputDataTablesCell into
            % outputDataBigTableCell
            outputDataBigTableCell.(string(tableType)).(string(currentSurvey)) = outputDataTablesCell.(string(tableType)).(string(currentSurvey)).(string(unique(extract(fieldnames(pageHeaders.(string(currentSurvey))), lettersPattern))));
            
            % - Specify spreadsheet formatting -
            format.(string(tableType)).(string(currentSurvey)).boldRows = 1;
            format.(string(tableType)).(string(currentSurvey)).bottomBorderRows = 1;
            format.(string(tableType)).(string(currentSurvey)).rightBorderRows = 2:height(outputDataBigTableCell.(string(tableType)).(string(currentSurvey)));
            if strcmp(tableType, 'percentPageCompleted')
                % calculate the cell fill colors for cells with numeric
                % data
                format.(string(tableType)).(string(currentSurvey)).isNumeric = cellfun(@isnumeric, outputDataBigTableCell.(string(tableType)).(string(currentSurvey))) & ~cellfun(@isempty, outputDataBigTableCell.(string(tableType)).(string(currentSurvey)));
                format.(string(tableType)).(string(currentSurvey)).fillColor = NaN(size(outputDataBigTableCell.(string(tableType)).(string(currentSurvey))));
                for rowNum = 1:height(format.(string(tableType)).(string(currentSurvey)).isNumeric)
                    for colNum = 2:width(format.(string(tableType)).(string(currentSurvey)).isNumeric)
                        if format.(string(tableType)).(string(currentSurvey)).isNumeric(rowNum, colNum)
                            % fill color where range is from 0-1
                            format.(string(tableType)).(string(currentSurvey)).fillColorAbsolute(rowNum, colNum) = excelInteriorColorConverter(colorMapCalc(outputDataBigTableCell.(string(tableType)).(string(currentSurvey)){rowNum, colNum} / 100, [255, 0, 0; 255, 255, 0; 0, 255, 0])); % cmap is [red, yellow, green]
                            % fill color where range is min(currentPhaseValues) - max(currentPhaseValues) 
                            % (i.e., map [0,1] onto [min(currentPhaseValues),max(currentPhaseValues)]
                            format.(string(tableType)).(string(currentSurvey)).fillColorRelative(rowNum, colNum) = excelInteriorColorConverter(colorMapCalc(interp1([min(cell2mat(outputDataBigTableCell.(string(tableType)).(string(currentSurvey))(rowNum, 2:end)) / 100), max(cell2mat(outputDataBigTableCell.(string(tableType)).(string(currentSurvey))(rowNum, 2:end)) / 100)], [0,1], outputDataBigTableCell.(string(tableType)).(string(currentSurvey)){rowNum, colNum} / 100), [255, 0, 0; 255, 255, 0; 0, 255, 0])); % cmap is [red, yellow, green]
                        end
                    end
                end
            end
        end
    end
end

%%
% --- Export data to spreadsheets ---
tableTypeList = {'percentPageCompleted_absColMap', 'percentPageCompleted_relColMap', 'pageCompleted'};
sheetCount = 2; % start at 2, because standard KPI (see above) will exist as sheet 1
for currentSurvey = {'Enrollment', 'Long', 'Short'}
    exportFileName = 'Dupuytrens KPI.xlsx';
    for tableTypeNum = 1:length(tableTypeList)
        % define the table name
        if contains(tableTypeList(tableTypeNum), '_')
            tableName = string(extractBefore(tableTypeList(tableTypeNum), '_'));
        else
            tableName = tableTypeList{tableTypeNum};
        end
        % write the data to a spreadsheet
        if strcmp(string(currentSurvey), 'Enrollment')
            writecell(outputDataBigTableCell.(tableName).(string(currentSurvey)), exportFileName, 'Sheet',strrep(sprintf('%s_Enroll', string(tableTypeList(tableTypeNum))), 'percent', '%'));
        else
            writecell(outputDataBigTableCell.(tableName).(string(currentSurvey)), exportFileName, 'Sheet',strrep(sprintf('%s_%s', string(tableTypeList(tableTypeNum)), string(currentSurvey)), 'percent', '%'));
        end
        % - format spreadsheets -
        e = actxserver('Excel.Application');
        try
            ewb = e.Workbooks.Open([pwd, '\', exportFileName]);
            e.DisplayAlerts = false;
            if contains(tableTypeList(tableTypeNum), 'percentPageCompleted')
                for colNum = 1:width(format.(tableName).(string(currentSurvey)).isNumeric)
                    for rowNum = 1:height(format.(tableName).(string(currentSurvey)).isNumeric)
                        if format.(tableName).(string(currentSurvey)).isNumeric(rowNum, colNum)
                            % highlight cells for color-coded magicalness
                            if contains(tableTypeList(tableTypeNum), 'absColMap')
                                % absolute color map
                                ewb.Worksheets.Item(sheetCount).Range(excelRangeFinder(rowNum, rowNum, colNum, colNum)).Interior.Color = format.(tableName).(string(currentSurvey)).fillColorAbsolute(rowNum, colNum);
                            else
                                % relative color map
                                ewb.Worksheets.Item(sheetCount).Range(excelRangeFinder(rowNum, rowNum, colNum, colNum)).Interior.Color = format.(tableName).(string(currentSurvey)).fillColorRelative(rowNum, colNum);
                            end
                        end
                    end
                end
            end
            % bold the title rows
            for rowNum = format.(tableName).(string(currentSurvey)).boldRows
                ewb.Worksheets.Item(sheetCount).Range(excelRangeFinder(rowNum, rowNum, 2, sum(~cellfun(@isempty, outputDataBigTableCell.(tableName).(string(currentSurvey))(rowNum, :))) + 1)).Font.Bold = 'True';
            end
            % bottom border of title rows
            for rowNum = format.(tableName).(string(currentSurvey)).bottomBorderRows
                ewb.Worksheets.Item(sheetCount).Range(excelRangeFinder(rowNum, rowNum, 2, sum(~cellfun(@isempty, outputDataBigTableCell.(tableName).(string(currentSurvey))(rowNum, :))) + 1)).Borders.Item('xlEdgeBottom').Weight = -4138;
            end
            % right border of first column in rows containing data
            for rowNum = format.(tableName).(string(currentSurvey)).rightBorderRows
                ewb.Worksheets.Item(sheetCount).Range(excelRangeFinder(rowNum, rowNum, 1, 1)).Borders.Item('xlEdgeRight').Weight = -4138;
            end
            ewb.Save;
            ewb.Close(false);
            e.Quit % Close excel application (otherwise, excel application won't be visible and it will need to be closed from the task manager)
        catch ME
            e.Quit % Close excel application (otherwise, excel application won't be visible and it will need to be closed from the task manager)
            rethrow(ME)
        end
        sheetCount = sheetCount + 1;
    end
end

% -------------------------------------------------------------------------
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






















