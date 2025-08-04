--USE Forward
--SELECT DISTINCT ph.Name--COUNT(*) as Count 
--FROM arc.dbo.Natalog n
--JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
--JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--WHERE diagnosis = 'PSOR' -- must have PSOR diagnosis --' ...
--    --AND ph.Name = 'Psoriasis Registry - Enrollment' -- must be enrolled in Psoriasis --' ...
--    --AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --


--USE Forward
--SELECT COUNT(*) as Count 
--FROM arc.dbo.Natalog n
--JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
--JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--WHERE diagnosis = 'PSOR' -- must have PSOR diagnosis --' ...
--    AND ph.Name = 'Psoriasis Registry - Enrollment' -- must be enrolled in Psoriasis --' ...
--	AND s.IsComplete = 1
--    --AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --

---- The number of questionnaires completed for the current phase ----
--USE Forward
--SELECT COUNT(*) as Count 
--FROM arc.dbo.Natalog n
--JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
--JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--WHERE diagnosis = 'DUP' -- must have PSOR diagnosis --' ...
--    --AND ph.Name Like '%Questionnaire (%Dupuytren%' -- must be enrolled in Psoriasis --' ...
--	--AND ph.Name Not Like '%Questionnaire (%Dupuytren%Short%'
--	AND ph.Name Like '%Questionnaire (Dupuytren)' -- must be enrolled in Psoriasis --' ...
--	AND ph.Active = 1
--	--AND s.IsComplete = 1
--	--AND up.StatusCode = ph.SentStatus
--	AND up.StatusCode = ph.CompleteStatus
--    --AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --

---- Modified Dupuytrens query that can be used to accurately extract aggregate KPI info ----
USE Forward 
--SELECT DISTINCT COUNT(*) AS 'Num Participants that Completed Page'
SELECT DISTINCT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode'
--SELECT TOP 10 q.Date
FROM Dupuytrens.dbo.Dup87_pg1 q -- grab info for specific page of specific questionnaire --' ...
JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Psoriasis.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
	--AND jul24 = ph.CompleteStatus
	AND ph.DatabasePrefix = 'Dup87' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
                                                                                     -- % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
                                                                                     -- % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}
ORDER BY up.UserId

--USE Forward
--SELECT COUNT(*) as Count 
--FROM arc.dbo.Natalog n
--JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
--JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId
--WHERE up.StatusCode = ph.SentStatus
--	--AND diagnosis = 'DUP' -- must have PSOR diagnosis --' ...
--    AND pj.Name = 'Dupuytrens' -- must be enrolled in Psoriasis --' ...
--	--AND ph.Name Not Like '%Questionnaire (%Dupuytren%Short%'
--	AND s.IsComplete = 1
--	--AND up.StatusCode = ph.SentStatus
--	--AND up.StatusCode = ph.CompleteStatus
--    --AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --

--USE Forward
--SELECT DISTINCT ph.Name, ph.DatabasePrefix--COUNT(*) as Count 
--FROM arc.dbo.Natalog n
--JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
--JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--WHERE diagnosis = 'PSOR' -- must have PSOR diagnosis --' ...
--	--AND ph.Active = 1 -- specify active surveys --
--    --AND ph.Name Like 'Psoriasis Registry -%onth Follow%' 
--	--AND ph.Name Like 'Monthly Questionnaire - February 2025' 
--	--AND s.IsComplete = 1
--    --AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

---- availablePhases: The list of unique survey phases and natalog phaseIDs in Psoriasis
--USE Forward
--SELECT nsr.ColumnTitle as 'NatalogField', DatabasePrefix, questid, StartDate, EndDate
--FROM Phase ph
--JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
--    AND ph.ProjectId = 46 -- 57 = PsO --' ...
----WHERE ph.DatabasePrefix like 'followup%' -- must have {SURVEYNAME} database prefix -- % SURVEYNAME = {'Enroll', 'monthly', 'followup'}

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

---- availableTables: The list of unique survey phases and pages in Psoriasis
--USE PsO_Registry
--SELECT TABLE_NAME
--FROM PsO_Registry.INFORMATION_SCHEMA.TABLES
--WHERE TABLE_NAME like 'followup%_pg%' -- must have {SURVEYNAME} in table name -- % SURVEYNAME = {'Enroll', 'monthly', 'followup'}

---- DUPUYTRENS availableTables: The list of unique survey phases and pages in Psoriasis
--USE Dupuytrens
--SELECT TABLE_NAME
--FROM Dupuytrens.INFORMATION_SCHEMA.TABLES
--WHERE TABLE_NAME like 'Dup%_pg%' -- must have {SURVEYNAME} in table name -- % SURVEYNAME = {'Enroll', 'Dup', 'Short'}

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

---- pageHeaders: The list of page numbers and page titles ALONG WITH the first survey that started using that specific order of page titles and numbers
--USE Forward
--SELECT DatabasePrefix, '[' + CAST(PageNumber as VARCHAR(3)) + '] ' + PageHeader as 'Page Header'
--FROM Page p
--JOIN Phase ph ON ph.SurveyToolId = p.SurveyToolId
--WHERE ProjectId = 57
--	AND ph.DatabasePrefix like 'monthly%'
--ORDER BY ph.SurveyToolId, DatabasePrefix, p.PageNumber

---- DUPUYTRENS pageHeaders: The list of page numbers and page titles ALONG WITH the first survey that started using that specific order of page titles and numbers
--USE Forward 
--SELECT DatabasePrefix, '[' + CAST(PageNumber as VARCHAR(3)) + '] ' + Pageheader as 'Page Header'
--FROM Page p
----JOIN SurveyTool st ON st.SurveyToolId = p.SurveyToolId -- match SurveyToolIDs between Forward.SurveyTools and Forward.Page --' ...
----JOIN Phase ph ON ph.PhaseId = st.PhaseId -- match phase IDs between Forward.Phase and Forward.SurveyTools --' ...
--JOIN Phase ph ON ph.SurveyToolId = p.SurveyToolId
--WHERE ph.DatabasePrefix like 'Dup%' -- must have {SURVEYNAME} database prefix --' ...
--	AND ProjectId = 46 -- specify the projectID... -- % SURVEYNAME = {'Enroll', 'Dup', 'Short'}
--ORDER BY ph.SurveyToolId, DatabasePrefix, p.PageNumber

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

---- questionnairesSent: The number of people who should have completed questionnaires (i.e., the number of questionnaires sent out)
--USE Forward
--SELECT COUNT(*) AS 'Questionnaires sent out', ph.DatabasePrefix
--FROM UserPhase up -- grab user activity --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'PsO_Registry' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	--AND ph.DatabasePrefix Like 'followup_%' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYSELECT = {'ph.Name = ''Psoriasis Registry - Enrollment''', 'ph.DatabasePrefix = ''monthly_YYYY03MM''', 'ph.DatabasePrefix Like ''followup_%'''} 
--                                                                                                                  -- % SURVEYDESCRIPTION = {'Psoriasis Registry - Enrollment PhaseId', 'Psoriasis Monthly Survey for month MM/YYYY', 'Psoriasis Followup Survey for patient'}
--	GROUP BY ph.DatabasePrefix
---- DUPUYTRENS questionnairesSent: The number of people who should have completed questionnaires (i.e., the number of questionnaires sent out)
--USE Forward
----SELECT COUNT(*) AS 'Questionnaires sent out', ph.DatabasePrefix
--SELECT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode'
--FROM UserPhase up -- grab user activity --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	AND ph.DatabasePrefix = 'Dup87' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
--                                                                                                  -- % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}
----GROUP BY ph.DatabasePrefix
--ORDER BY up.UserId

--USE Forward SELECT Top 10 * FROM UserPhase --ORDER BY GUID
--USE Forward SELECT *--GUID, Date 
--FROM Dupuytrens.dbo.Dup87_pg1 ORDER BY GUID

--USE Forward SELECT *--GUID, Date 
--FROM arc.dbo.Natalog WHERE GUID Like '2F7B08AF-%' ORDER BY GUID

--USE Forward
----SELECT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode'
--SELECT COUNT(*) AS 'Questionnaires sent out'
--FROM UserPhase up -- grab user activity --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	AND ph.DatabasePrefix like 'Dup%' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
--                                                                                                  -- % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}
--	--GROUP BY ph.DatabasePrefix

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

---- pageStarted: The number of people who started the page
--USE Forward 
--SELECT COUNT(*) AS 'Num Participants that Started Page'
--FROM PsO_Registry.dbo.followup_20250107_pg1 q -- grab info for specific page of specific questionnaire --' ...
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Psoriasis.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'PsO_Registry' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	AND ph.DatabasePrefix = 'followup_20250107' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'Enrollment', 'monthly_YYYY03MM', 'followup_%'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug', 'followup_pg8' NOT 'followup_6_mo_pg8'))
--																								  -- % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''monthly_YYYY03MM''', 'ph.DatabasePrefix = ''followup_20250107''' OR 'ph.DatabasePrefix = ''followup_6_mo'''} 
--																						 		  -- % SURVEYDESCRIPTION = {'Psoriasis Registry - Enrollment PhaseId', 'Psoriasis Monthly Survey for MM/YYYY', 'Psoriasis Followup Survey for version "20250107" or "6_mo"'}

---- DUPUYTRENS pageStarted: The number of people who started the page
--USE Forward 
--SELECT COUNT(*) AS 'Num Participants that Started Page'
--FROM Dupuytrens.dbo.Dup86_pg1 q -- grab info for specific page of specific questionnaire --' ...
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Psoriasis.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	AND ph.DatabasePrefix = 'Dup86' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
--                                                                                     -- % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
--                                                                                     -- % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}

--USE Forward 
--SELECT up.ActivityDate AS 'ActivityDate' --COUNT(*) AS 'Num Participants that Started Page'
--FROM Dupuytrens.dbo.Dup86_pg1 q -- grab info for specific page of specific questionnaire --' ...
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Psoriasis.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	AND ph.DatabasePrefix = 'Dup86' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --'
--	AND jan24 = ph.StartedStatus --]; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
--                                                                                     -- % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
--                                                                                     -- % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

---- pageCompleted: The number of people who completed the page
--USE Forward 
--SELECT COUNT(*) AS 'Num Participants that Completed Page'
--FROM PsO_Registry.dbo.followup_20250107_pg1 q -- grab info for specific page of specific questionnaire --' ...
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Psoriasis.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'PsO_Registry' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	AND Complete = 1 -- only include complete pages --
--	AND jul24 = ph.CompleteStatus
--	AND ph.DatabasePrefix = 'followup_20250107' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'Enrollment', 'monthly_YYYY03MM', 'followup_%'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug', 'followup_pg8' NOT 'followup_6_mo_pg8'))
--																								              -- % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''monthly_YYYY03MM''', 'ph.DatabasePrefix = ''followup_20250107''' OR 'ph.DatabasePrefix = ''followup_6_mo'''} 
--																						 		              -- % SURVEYDESCRIPTION = {'Psoriasis Registry - Enrollment PhaseId', 'Psoriasis Monthly Survey for MM/YYYY', 'Psoriasis Followup Survey for version "20250107" or "6_mo"'}

---- DUPUYTRENS pageCompleted: The number of people who completed the page
--USE Forward 
--SELECT COUNT(*) AS 'Num Participants that Completed Page'
----SELECT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode'
--FROM Dupuytrens.dbo.Dup86_pg1 q -- grab info for specific page of specific questionnaire --' ...
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Psoriasis.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Psoriasis Registry (aka Project) --' ...
--	AND Complete = 1 -- only include complete pages --
--	AND jan24 = ph.CompleteStatus
--	AND ph.DatabasePrefix = 'Dup86' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --']; % SURVEYPAGEHEADER = sprintf('%s_pg%d', {'EnrollDup', 'Dup##', 'Short##'}, pageNum (EXCEPT ONE PAGE IS 'pgDrug'))
--                                                                                     -- % SURVEYSELECT = {'ph.Name LIKE ''%Enrollment%''', 'ph.DatabasePrefix = ''Dup##''', 'ph.DatabasePrefix = ''Short##'''} 
--                                                                                     -- % SURVEYDESCRIPTION = {'Dupuytrens Enrollment PhaseId', 'Dupuytrens Long Survey for ph##', 'Dupuytrens Short Survey for ph##'}

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

---- Gives the number of surveys per each day of the week
--USE Forward
--SELECT COUNT(DATEPART(weekday, up.ActivityDate)) AS 'Count per Weekday'
--FROM arc.dbo.Natalog n
--JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --' ...
--JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--WHERE diagnosis = 'PSOR' -- must have PSOR diagnosis --' ...
--    --AND ph.Name = 'Psoriasis Registry - Enrollment' -- must be enrolled in Psoriasis --' ...
--	--AND s.IsComplete = 1
--    --AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --
--GROUP BY DATEPART(weekday, up.ActivityDate)

--USE Forward;
--WITH WeekDays AS (
--SELECT 
--	CASE 
--		WHEN datepart(weekday, up.ActivityDate) = 1 THEN 'Sunday'
--		WHEN datepart(weekday, up.ActivityDate) = 2 THEN 'Monday'
--		WHEN datepart(weekday, up.ActivityDate) = 3 THEN 'Tuesday'
--		WHEN datepart(weekday, up.ActivityDate) = 4 THEN 'Wednesday'
--		WHEN datepart(weekday, up.ActivityDate) = 5 THEN 'Thursday'
--		WHEN datepart(weekday, up.ActivityDate) = 6 THEN 'Friday'
--		WHEN datepart(weekday, up.ActivityDate) = 7 THEN 'Saturday'
--	END as 'Weekday'
--	FROM
--arc.dbo.Natalog n
--	JOIN UserPhase up ON up.UserId = n.GUID
--	JOIN Status s on s.Code = up.StatusCode
--	JOIN Phase ph ON ph.PhaseId = up.PhaseId
--	WHERE diagnosis = 'PSOR' )

--SELECT Weekday, COUNT(*)
--	FROM WeekDays
--	GROUP BY Weekday



--USE Forward
--SELECT DATEPART(weekday, up.ActivityDate) AS 'Weekday Name', COUNT(DATEPART(weekday, up.ActivityDate)) AS 'Count'
--FROM Dupuytrens.dbo.Dup88_pg1 q -- grab info for specific page of specific questionnaire --
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.Short88_pg18 and Forward.UserPhase --
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
--	AND ph.DatabasePrefix = 'Dup88' -- limit UserPhase to only include the Dupuytrens Short Survey for ph88 --
--	AND Complete = 1 -- only include complete pages --
--	AND jan25 = ph.CompleteStatus -- confirm CompleteStatus --
--GROUP BY DATEPART(weekday, up.ActivityDate)


--USE Forward;
--WITH WeekDays AS (
--SELECT 
--	CASE 
--		WHEN datepart(weekday, up.ActivityDate) = 1 THEN 'Sunday'
--		WHEN datepart(weekday, up.ActivityDate) = 2 THEN 'Monday'
--		WHEN datepart(weekday, up.ActivityDate) = 3 THEN 'Tuesday'
--		WHEN datepart(weekday, up.ActivityDate) = 4 THEN 'Wednesday'
--		WHEN datepart(weekday, up.ActivityDate) = 5 THEN 'Thursday'
--		WHEN datepart(weekday, up.ActivityDate) = 6 THEN 'Friday'
--		WHEN datepart(weekday, up.ActivityDate) = 7 THEN 'Saturday'
--	END as 'Weekday'
--	FROM
--Dupuytrens.dbo.Dup88_pg1 q -- grab info for specific page of specific questionnaire --
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.Short88_pg18 and Forward.UserPhase --
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
--	AND ph.DatabasePrefix = 'Dup88' -- limit UserPhase to only include the Dupuytrens Short Survey for ph88 --
--	AND Complete = 1 -- only include complete pages --
--	AND jan25 = ph.CompleteStatus -- confirm CompleteStatus --
--	)

--SELECT Weekday, COUNT(*)
--	FROM WeekDays
--	GROUP BY Weekday



--USE Forward;
--WITH PostalCodes AS (
--SELECT 
--	CASE 
--		--WHEN n.zip Like '#####' THEN n.zip --CAST(n.zip AS numeric)
--		WHEN n.zip Not Like '_____' THEN 00000
--		ELSE SUBSTRING(n.zip, 1, 3)
--	END as 'PostalCode'
----SELECT n.zip AS 'Postal Code', up.ActivityDate AS 'Date'   
----SELECT n.zip AS 'Postal Code'
--FROM Dupuytrens.dbo.Dup88_pg1 q -- grab info for specific page of specific questionnaire --
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.Short88_pg18 and Forward.UserPhase --
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
--	AND ph.DatabasePrefix = 'Dup88' -- limit UserPhase to only include the Dupuytrens Short Survey for ph88 --
--	AND Complete = 1 -- only include complete pages --
--	AND jan25 = ph.CompleteStatus -- confirm CompleteStatus --
----GROUP BY DATEPART(weekday, up.ActivityDate)
--	)

--SELECT PostalCode, COUNT(*)
--	FROM PostalCodes
--	GROUP BY PostalCode



------ Get a list of dynamic project variables (e.g., 'DUP', 'Psoriasis - Enrollment Survey', etc.) ------
--USE Forward
--SELECT *
--FROM Forward.dbo.ProjectCategory pc
--JOIN Forward.dbo.ProjectCategories pcs on pc.ProjectCategoryId = pcs.ProjectCategoryId 
--JOIN Forward.dbo.Project p on pcs.ProjectId = p.ProjectId -- join with Forward.dbo.ProjectCategories, Forward.dbo.Project, Forward.dbo.SurveyTool 

select * from Forward.dbo.ProjectCategory
select * from Forward.dbo.ProjectCategories
select * from Forward.dbo.ProjectDiagnosis
select * from Forward.dbo.Project

--select distinct s.Name, pj.Name
--	from SurveyTool s
--	join Phase ph on ph.SurveyToolId = s.SurveyToolId
--	join Project pj on pj.ProjectId = ph.ProjectId
--	join ProjectCategories pc on ph.ProjectId = pc.ProjectId
--	join ProjectCategory py on py.ProjectCategoryId = pc.ProjectCategoryId
--	where 
--		--py.ProjectCategoryId = 5 -- OA
--		py.ProjectCategoryId = 16 -- PSOR

USE Forward
SELECT p.Name, pcc.Code
FROM Project p
JOIN ProjectCategories pc on p.ProjectId = pc.ProjectId
JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId
	--WHERE p.ProjectId = 14

-------------------------------------------------------------------------------------------
-- DEPRICIATED: Create availableDatabasePrefixes --
USE Forward
SELECT p.Name as 'Project Name', pcc.Code as 'Diagnosis Code', ph.DatabasePrefix, ph.questid as 'QuestID', ph.PhaseId as 'PhaseID', nsr.ColumnTitle as 'NatalogField', ph.StartDate, ph.EndDate
FROM Project p
JOIN ProjectCategories pc on p.ProjectId = pc.ProjectId
JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId
JOIN Phase ph on ph.ProjectId = p.ProjectId
JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
WHERE p.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'Dupuytrens') -- where {PROJECTNAME} is the project name (e.g., NDB, Dupuytrens, PsO_Registry, etc.) --
ORDER BY DatabasePrefix ASC

------------------------------------------------
-- DEPRICIATED: Dynamically replace availablePhases --
USE Forward
SELECT nsr.ColumnTitle as 'NatalogField', DatabasePrefix, questid, StartDate, EndDate
FROM Phase ph 
JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
    AND ph.ProjectId = 46 -- 46 = Dup --
WHERE ph.DatabasePrefix like 'Dup%' -- WHERE ph.DatabasePrefix like '{SURVEYNAME}%' -- SURVEYNAME = {'Enroll', 'Dup', 'Short'}

------------------------------------------------
-- DEPRICIATED: Dynamically replace availableTables --
USE Dupuytrens
SELECT *
FROM Dupuytrens.INFORMATION_SCHEMA.TABLES
--ORDER BY TABLE_NAME
WHERE TABLE_NAME like 'Dup%_pg%'

-- aka availableTablesV2
USE Dupuytrens
SELECT TABLE_NAME, TABLE_TYPE
FROM Dupuytrens.INFORMATION_SCHEMA.TABLES
ORDER BY TABLE_NAME

------------------------------------------------
-- DEPRICIATED: Dynamically replace pageHeaders --
USE Forward
SELECT DatabasePrefix, '[' + CAST(PageNumber as VARCHAR(3)) + '] ' + Pageheader as 'Page Header' 
FROM Page p 
JOIN Phase ph ON ph.SurveyToolId = p.SurveyToolId -- match SurveyToolIDs between Forward.Phase and Forward.Page --
WHERE ph.DatabasePrefix like 'Dup%' -- must have {SURVEYNAME} database prefix --
	AND ProjectId = 46 -- specify the projectID... --
ORDER BY ph.SurveyToolId, DatabasePrefix, p.PageNumber

------------------------------------------------
-- Create projectDatabaseInfo table -- 
-- this table returns table and phase info for a specific project --
SELECT DISTINCT 
	pj.Name as 'ProjectName', pcc.Code as 'DiagnosisCode', ph.DatabasePrefix, 
    pj.Name + '.dbo.' + ph.DatabasePrefix + '_' + 
    CASE 
        WHEN s.SectionTypeId = 5 THEN 'pgDrug'
        ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3))
    END AS 'PageTable',
    ph.DatabasePrefix + '_' + CASE 
        WHEN s.SectionTypeId = 5 THEN 'pgDrug'
        ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3))
    END as 'TableName', '[' + CAST(p.PageNumber as VARCHAR(3)) + '] ' + Pageheader as 'PageHeader', p.PageNumber, ph.questid, ph.PhaseId as 'PhaseID', nsr.ColumnTitle as 'PhaseStatus', ph.Name as 'PhaseName', ph.StartDate, ph.EndDate, ph.InviteStatus, ph.SentStatus, ph.StartedStatus, ph.CompleteStatus
    FROM Forward.dbo.Phase ph
    JOIN Forward.dbo.Project pj ON pj.ProjectId = ph.ProjectId
    JOIN Forward.dbo.Page p ON p.SurveyToolId = ph.SurveyToolId
    JOIN Forward.dbo.Section s ON s.PageId = p.PageId
	JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId
	JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId
	JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
    WHERE pj.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'NDB') -- specify the projectID... --
	ORDER BY PhaseStatus


USE Forward
SELECT * 
--FROM Phase ph
--JOIN SurveyTool st on ph.SurveyToolId = st.SurveyToolId
FROM Project pj
JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId
JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId
	   
------------------------------------------------
-- Dynamically replace questionnairesSent --
USE Forward
SELECT DISTINCT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode'
FROM UserPhase up -- grab user activity --
JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
	AND ph.DatabasePrefix = 'Dup80'

-- questionnairesSentV2
USE Forward
SELECT DISTINCT up.UserId AS 'UserID', n.zip AS 'PostalCode', n.state AS 'StateAbbr', n.country AS 'CountryCode', ph.DatabasePrefix
FROM UserPhase up -- grab user activity --
JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --

------------------------------------------------
-- DEPRICIATED: Dynamically replace pageCompleted --
USE Forward
SELECT DISTINCT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode'
FROM Dupuytrens.dbo.{SURVEYPAGEHEADER} q -- grab info for specific page of specific questionnaire --
JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --
JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
	AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --
    AND {NATALOGPHASEID} = ph.CompleteStatus

-- DEPRICIATED: pageCompletedV2 (aka pageData)
USE Forward
--SELECT DISTINCT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode', ph.DatabasePrefix, jan21 as 'SurveyStatus'
SELECT TOP 100 *
FROM Dupuytrens.dbo.Dup80_pg1 q -- grab info for specific page of specific questionnaire --
JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --
JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
	--AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --
    --AND jan21 = ph.CompleteStatus
ORDER BY up.UserId 



-- try to generate a query that returns pageData as described in the Matlab script
-- NOTE: I can query pageCompletedV2 (minus the Dup80_pg## in the from line) to get a list of who 
--       was sent which surveys, BUT I still need to query each 'Dup80_pg##' table to get a list of
--       who STARTED and WHAT TIME THEY STARTED each page. (NOTE: 'Date' field of PROJECTNAME_pg## table IS DEFINITELY more accurate than 'ActivityDate' field of UserPhase table
-- STEPS TO APPROACH THIS:
--     IDEA 1:
--         1) Start with questionnairesSentV2
--         2) Add *recd columns from Natalog
--         3) Add Date column from each page num
--     IDEA 2: DOES NOT WORK CUS STEP 4 POPULATES EACH 'Date' FIELD ENTRY FOR THE SAME PATIENT AS THE DATE THEY STARTED THAT SPECIFIC PROJECTNAME_pg##
--         1) Start with questionnairesSentV2
--         2) Separately query all *recd columns from Natalog
--         3) Combine *recd data with questionnairesSentV2 data in Matlab
--         4) Separately query PROJECTNAME_pg## data for survey with most pages within project (since querying XXX_pg1 (without filtering for DatabasePrefix) returns all patients that completed a pg1 for all DatabasePrefixes)
--     IDEA 3:
--         1) Start with questionnairesSentV2
--         2) Separately query all *recd columns from Natalog
--         3) Combine *recd data with questionnairesSentV2 data in Matlab
--         4a) Separately query PROJECTNAME_pg## data for EACH page of EACH survey for EACH phase
--		   4b) As PROJECTNAME_pg## tables are queried, update the 'pgStartDate' column for the current phase/survey for all patients  
--     IDEA 4:
--         1) Start with projectDatabaseInfo, only with DatabasePrefix, PageTable, PageNumber, and PhaseStatus
--         2) Loop through each table name, appending Date (from page table), Phase Status (e.g., jan21 from Natalog table), and PostalCode (zip field from Natalog)


USE Forward;
DROP TABLE IF EXISTS #pageQueryResults
CREATE TABLE #pageQueryResults (
	   DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT
);

DROP TABLE IF EXISTS #pageDataFull
CREATE TABLE #pageDataFull (
	   DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT
);

DECLARE @currentTableName varchar(max) = '';
DECLARE @currentTableNameShort varchar(max) = '';
DECLARE @currentPageNum INT;
DECLARE @currentPhaseStatus varchar(max) = '';
DECLARE @currentDatabasePrefix varchar(max) = '';
DECLARE @sql varchar(max) = '';
-- Generate a list of page table names, their page numbers, their phase statuses, and their database prefixes for the current project. These will be used to dynamically query page tables. --
DECLARE tableListCursor CURSOR FOR SELECT DISTINCT 
	pj.Name + '.dbo.' + ph.DatabasePrefix + '_' + 
    CASE 
        WHEN s.SectionTypeId = 5 THEN 'pgDrug'
        ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3))
    END AS 'PageTable', 
	ph.DatabasePrefix + '_' + CASE 
        WHEN s.SectionTypeId = 5 THEN 'pgDrug'
        ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3))
    END as 'TableName', ph.DatabasePrefix, p.PageNumber, nsr.ColumnTitle as 'PhaseStatus'
    FROM Forward.dbo.Phase ph
    JOIN Forward.dbo.Project pj ON pj.ProjectId = ph.ProjectId
    JOIN Forward.dbo.Page p ON p.SurveyToolId = ph.SurveyToolId
    JOIN Forward.dbo.Section s ON s.PageId = p.PageId
	JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId
	JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId
	JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
    WHERE pj.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'Dupuytrens') --VARIABLE--
		--AND pcc.Code = 'DUP' --VARIABLE--
	ORDER BY p.PageNumber;
OPEN tableListCursor;
FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus;
WHILE @@FETCH_STATUS = 0
BEGIN
	-- pageQuery into #pageQueryResults: Retrieve all the rows for the current page table, which contain UserID (from page table), LastActivityDate (from page table), SurveyStatus (from Natalog via e.g. "n.jan21"), and PostalCode (from Natalog) --  
	SET @sql = N'
		WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', n.zip as ' + QUOTENAME('PostalCode') + ', n.state as ' + QUOTENAME('StateAbbr') + ', n.country as ' + QUOTENAME('CountryCode') + ' 
						   FROM ' + @currentTableName + ' q
						   JOIN arc.dbo.Natalog n ON q.guid = n.guid) 

		INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode)
		SELECT UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode
		FROM pageQuery';
	EXEC(@sql);
	--PRINT(@sql);

	-- pageDataFullQuery: Add @currentDatabasePrefix and either "pg0#" or "pg##" to the page table results from #pageQueryResults --  
	WITH pageDataFullQuery AS (SELECT @currentDatabasePrefix AS 'DatabasePrefixVar', 
							 		  'pg' + CASE WHEN @currentPageNum > 9 THEN CAST(@currentPageNum AS VARCHAR(3)) ELSE '0' + CAST(@currentPageNum AS VARCHAR(3)) END as 'PageNumber', 
									  pqr.UserID, pqr.LastActivityDate, pqr.SurveyStatus, pqr.PostalCode, pqr.StateAbbr, pqr.CountryCode FROM #pageQueryResults pqr)

	-- pageDataFullQuery into #pageDataFull -- 
	INSERT INTO #pageDataFull (DatabasePrefix, PageNum, UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode)
	SELECT *
	FROM pageDataFullQuery

	-- Clear #pageQueryResults, so #pageQueryResults does not get appended to itself resulting in exponential duplication of its data --
	DELETE FROM #pageQueryResults

	FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus;
END
CLOSE tableListCursor;
DEALLOCATE tableListCursor;
GO

SELECT * 
FROM #pageDataFull
ORDER BY UserID, PageNum


-- Create a table that condenses the line-level data from #pageDataFull by pivoting SurveyStatus and LastActivityDate into columns by page -- 
DROP TABLE IF EXISTS #pageDataFinal
CREATE TABLE #pageDataFinal (
	   UserID VARCHAR(MAX), PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT, DatabasePrefix VARCHAR(MAX), pg01_SurveyStatus INT, pg01_LastActivityDate DATETIME, pg02_SurveyStatus INT, pg02_LastActivityDate DATETIME, pg03_SurveyStatus INT, pg03_LastActivityDate DATETIME, pg04_SurveyStatus INT, pg04_LastActivityDate DATETIME, pg05_SurveyStatus INT, pg05_LastActivityDate DATETIME, pg06_SurveyStatus INT, pg06_LastActivityDate DATETIME, pg07_SurveyStatus INT, pg07_LastActivityDate DATETIME, pg08_SurveyStatus INT, pg08_LastActivityDate DATETIME, pg09_SurveyStatus INT, pg09_LastActivityDate DATETIME, pg10_SurveyStatus INT, pg10_LastActivityDate DATETIME, pg11_SurveyStatus INT, pg11_LastActivityDate DATETIME, pg12_SurveyStatus INT, pg12_LastActivityDate DATETIME, pg13_SurveyStatus INT, pg13_LastActivityDate DATETIME, pg14_SurveyStatus INT, pg14_LastActivityDate DATETIME, pg15_SurveyStatus INT, pg15_LastActivityDate DATETIME, pg16_SurveyStatus INT, pg16_LastActivityDate DATETIME, pg17_SurveyStatus INT, pg17_LastActivityDate DATETIME, pg18_SurveyStatus INT, pg18_LastActivityDate DATETIME, pg19_SurveyStatus INT, pg19_LastActivityDate DATETIME, pg20_SurveyStatus INT, pg20_LastActivityDate DATETIME, pg21_SurveyStatus INT, pg21_LastActivityDate DATETIME, pg22_SurveyStatus INT, pg22_LastActivityDate DATETIME, pg23_SurveyStatus INT, pg23_LastActivityDate DATETIME, pg24_SurveyStatus INT, pg24_LastActivityDate DATETIME
);

INSERT INTO #pageDataFinal
SELECT pdfss.UserID, pdfss.PostalCode, pdfss.StateAbbr, pdfss.CountryCode, pdfss.DatabasePrefix, 
	   pdfss.pg01 AS 'pg01_SurveyStatus', pdflad.pg01 AS 'pg01_LastActivityDate', pdfss.pg02 AS 'pg02_SurveyStatus', pdflad.pg02 AS 'pg02_LastActivityDate', pdfss.pg03 AS 'pg03_SurveyStatus', pdflad.pg03 AS 'pg03_LastActivityDate', pdfss.pg04 AS 'pg04_SurveyStatus', pdflad.pg04 AS 'pg04_LastActivityDate', pdfss.pg05 AS 'pg05_SurveyStatus', pdflad.pg05 AS 'pg05_LastActivityDate', pdfss.pg06 AS 'pg06_SurveyStatus', pdflad.pg06 AS 'pg06_LastActivityDate', pdfss.pg07 AS 'pg07_SurveyStatus', pdflad.pg07 AS 'pg07_LastActivityDate', pdfss.pg08 AS 'pg08_SurveyStatus', pdflad.pg08 AS 'pg08_LastActivityDate', pdfss.pg09 AS 'pg09_SurveyStatus', pdflad.pg09 AS 'pg09_LastActivityDate', pdfss.pg10 AS 'pg10_SurveyStatus', pdflad.pg10 AS 'pg10_LastActivityDate', pdfss.pg11 AS 'pg11_SurveyStatus', pdflad.pg11 AS 'pg11_LastActivityDate', pdfss.pg12 AS 'pg12_SurveyStatus', pdflad.pg12 AS 'pg12_LastActivityDate', pdfss.pg13 AS 'pg13_SurveyStatus', pdflad.pg13 AS 'pg13_LastActivityDate', pdfss.pg14 AS 'pg14_SurveyStatus', pdflad.pg14 AS 'pg14_LastActivityDate', pdfss.pg15 AS 'pg15_SurveyStatus', pdflad.pg15 AS 'pg15_LastActivityDate', pdfss.pg16 AS 'pg16_SurveyStatus', pdflad.pg16 AS 'pg16_LastActivityDate', pdfss.pg17 AS 'pg17_SurveyStatus', pdflad.pg17 AS 'pg17_LastActivityDate', pdfss.pg18 AS 'pg18_SurveyStatus', pdflad.pg18 AS 'pg18_LastActivityDate', pdfss.pg19 AS 'pg19_SurveyStatus', pdflad.pg19 AS 'pg19_LastActivityDate', pdfss.pg20 AS 'pg20_SurveyStatus', pdflad.pg20 AS 'pg20_LastActivityDate', pdfss.pg21 AS 'pg21_SurveyStatus', pdflad.pg21 AS 'pg21_LastActivityDate', pdfss.pg22 AS 'pg22_SurveyStatus', pdflad.pg22 AS 'pg22_LastActivityDate', pdfss.pg23 AS 'pg23_SurveyStatus', pdflad.pg23 AS 'pg23_LastActivityDate', pdfss.pg24 AS 'pg24_SurveyStatus', pdflad.pg24 AS 'pg24_LastActivityDate'
FROM (  -- Create a table that pivots on the survey status (i.e., columns include DatabasePrefix, UserID, PostalCode, pg01, pg02, pg03, ... where pg## stores the SurveyStatus field for pg##) --
	    SELECT DatabasePrefix, UserID, MAX(PostalCode) AS 'PostalCode', MAX(StateAbbr) AS 'StateAbbr', MAX(CountryCode) AS 'CountryCode', 
		--MAX(pg01) AS 'pg01', MAX(pg02) AS 'pg02', MAX(pg03) AS 'pg03', MAX(pg04) AS 'pg04', MAX(pg05) AS 'pg05', MAX(pg06) AS 'pg06', MAX(pg07) AS 'pg07', MAX(pg08) AS 'pg08', MAX(pg09) AS 'pg09', MAX(pg10) AS 'pg10', MAX(pg11) AS 'pg11', MAX(pg12) AS 'pg12', MAX(pg13) AS 'pg13', MAX(pg14) AS 'pg14', MAX(pg15) AS 'pg15', MAX(pg16) AS 'pg16', MAX(pg17) AS 'pg17', MAX(pg18) AS 'pg18', MAX(pg19) AS 'pg19', MAX(pg20) AS 'pg20', MAX(pg21) AS 'pg21', MAX(pg22) AS 'pg22', MAX(pg23) AS 'pg23', MAX(pg24) AS 'pg24'
		MAX(pg01) AS 'pg01', MAX(pg02) AS 'pg02', MAX(pg03) AS 'pg03', MAX(pg04) AS 'pg04', MAX(pg05) AS 'pg05', MAX(pg06) AS 'pg06', MAX(pg07) AS 'pg07', MAX(pg08) AS 'pg08', MAX(pg09) AS 'pg09', MAX(pg10) AS 'pg10', MAX(pg11) AS 'pg11', MAX(pg12) AS 'pg12', MAX(pg13) AS 'pg13', MAX(pg14) AS 'pg14', MAX(pg15) AS 'pg15', MAX(pg16) AS 'pg16', MAX(pg17) AS 'pg17', MAX(pg18) AS 'pg18', MAX(pg19) AS 'pg19', MAX(pg20) AS 'pg20', MAX(pg21) AS 'pg21', MAX(pg22) AS 'pg22', MAX(pg23) AS 'pg23', MAX(pg24) AS 'pg24'
		FROM (
				SELECT *
				FROM #pageDataFull
				PIVOT (
					MAX(SurveyStatus)
					FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)
					) AS pageDataFullPivot
				) pdfss_pre
		GROUP BY DatabasePrefix, UserID
	) pdfss
JOIN ( -- Create a table that pivots on the last activity date (i.e., columns include DatabasePrefix, UserID, PostalCode, pg01, pg02, pg03, ... where pg## stores the LastActivityDate field for pg##) --
	  SELECT *
  	  FROM #pageDataFull
	  PIVOT (
	  MAX(LastActivityDate)
	  FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)
	  ) AS pageDataFullPivot) pdflad ON pdfss.UserID = pdflad.UserID 
									AND pdfss.DatabasePrefix = pdflad.DatabasePrefix
ORDER BY UserID, DatabasePrefix


SELECT *
FROM #pageDataFinal

--DECLARE @comparisonDate AS DATE = CAST('2025-07-05 23:27:58.997' as DATE)
DECLARE @comparisonDate AS DATE = GETDATE();
SELECT pdf.*, (CASE WHEN DATEDIFF (DAY, up.ActivityDate, @comparisonDate) <= 7 AND up.StatusCode LIKE '%9' THEN 1 ELSE 0 END) AS SurveyStartedWithinWeek
    FROM #pageDataFinal pdf
    JOIN Phase ph ON ph.DatabasePrefix = pdf.DatabasePrefix 
    JOIN UserPhase up ON ph.PhaseId = up.PhaseId AND pdf.UserID = up.UserId
	JOIN AspNetUsers a ON a.Id = up.UserId
	ORDER BY SurveyStartedWithinWeek
    WHERE DATEDIFF (DAY, up.ActivityDate, @comparisonDate) <= 7
        AND a.ProjectId = 46
        AND up.StatusCode LIKE '%9'
    ORDER BY up.ActivityDate DESC

INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0','Text;Database=C:\Users\Matt\Documents\GitHub\Dupuytren-Scripts\SQL Export Data\Dupuytrens SQL Export Test 02.csv;HDR=YES;FMT=Delimited','SELECT * FROM [Sheet1$]')
SELECT * FROM #pageDataFinal

SELECT DISTINCT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode', ph.DatabasePrefix FROM UserPhase up JOIN Phase ph ON ph.PhaseId = up.PhaseId JOIN Project pj ON pj.ProjectId = ph.ProjectId JOIN arc.dbo.Natalog n ON n.guid = up.UserId WHERE up.StatusCode = ph.SentStatus     AND pj.Name = 'NDB')

---- WORKS --
--USE Forward
--DROP TABLE IF EXISTS #temp
--CREATE TABLE #temp (
--       UserID UNIQUEIDENTIFIER, PageStartDate datetime, SurveyStatus INT, PageNum INT
--)

--DECLARE @currentPageNum INT = 1;
--DECLARE @totalPageNum INT = 24; --(SELECT COUNT(*) FROM #temp);

--WHILE @currentPageNum <= @totalPageNum
--BEGIN
--	--SELECT DISTINCT up.UserId AS 'UserID', up.ActivityDate AS 'ActivityDate', n.zip AS 'PostalCode', ph.DatabasePrefix, jan21 as 'SurveyStatus'
--	--SELECT DISTINCT TOP 100 *
--	WITH tempQuery AS (SELECT up.UserId AS 'UserID', q.Date as 'PageStartDate', jan21 as 'SurveyStatus', @currentPageNum as 'PageNumber'
--	FROM Dupuytrens.dbo.Dup80_pg20 q -- grab info for specific page of specific questionnaire --
--	JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --
--	JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
--	JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
--	--JOIN Page p ON p.SurveyToolId = ph.SurveyToolId
--	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
--	JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
--	--JOIN Forward.dbo.Section s ON s.PageId = p.PageId
--	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
--		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
--		--AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --
--		--AND jan21 = ph.CompleteStatus
--		AND DatabasePrefix = 'Dup80')

--	INSERT INTO #temp (UserID, PageStartDate, SurveyStatus, PageNum)
--	SELECT UserID, PageStartDate, SurveyStatus, PageNumber
--	FROM tempQuery

--	SET @currentPageNum = @currentPageNum + 1
--END

--SELECT TOP 1000 * 
--FROM #temp
--ORDER BY UserID


---- WORKS --
--USE Forward
--DROP TABLE IF EXISTS #temp
--CREATE TABLE #temp (
--       GUID UNIQUEIDENTIFIER, Date datetime
--)

--DECLARE @TablenameList VARCHAR(MAX) = 'Dupuytrens.dbo.Dup80_pg1, Dupuytrens.dbo.Dup80_pg2, Dupuytrens.dbo.Dup80_pg3';
--DECLARE @sql varchar(max) = '';
--DECLARE @Tablename varchar(max) = '';
--DECLARE Table_Cursor CURSOR FOR
--SELECT value AS Tablename FROM STRING_SPLIT(@TablenameList, ',');
--OPEN Table_Cursor;
--FETCH NEXT FROM Table_Cursor INTO @Tablename;
--WHILE @@FETCH_STATUS = 0
--BEGIN
--    SET @sql = '
--		WITH tempQuery AS (
--			SELECT GUID, Date 
--			FROM ' + @Tablename + ' q
--			)
--        INSERT INTO #temp  
--        SELECT GUID, Date
--        FROM tempQuery';
--    EXEC(@sql);

--    FETCH NEXT FROM Table_Cursor INTO @Tablename;
--END
--CLOSE Table_Cursor;
--DEALLOCATE Table_Cursor;
--GO

--SELECT TOP 1000 * 
--FROM #temp
--ORDER BY GUID


SELECT q.GUID as 'UserID', q.Date as 'LastActivityDate', n.zip as 'PostalCode', CASE 
        WHEN country != 0 THEN 'International'
		WHEN n.state IN ('AK','AL','AR','AS','AZ','CA','CO','CT','DC','DE','FL','FM','GA','GU','HI','IA','ID','IL','IN','KS','KY','LA','MA','MD','ME','MH','MI','MN','MO','MP','MS','MT','NC','ND','NE','NH','NJ','NM','NV','NY','OH','OK','OR','PA','PR','PW','RI','SC','SD','TN','TX','UT','VA','VI','VT','WA','WI','WV','WY') THEN n.state
        WHEN (n.zip like ('[0-9][0-9][0-9][0-9][0-9]')) OR (n.zip like ('[0-9][0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]')) THEN n.zip
		ELSE 'International'
    END AS 'PostalCodeState', n.state, n.country
FROM Dupuytrens.dbo.Dup80_pg11 q
JOIN arc.dbo.Natalog n ON q.guid = n.guid

USE arc;
SELECT *
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Natalog'


USE arc;
SELECT n.GUID, dob, sex, zip, n.state, country 
FROM arc.dbo.Natalog n
--GROUP BY n.GUID

USE Forward;
SELECT TOP 100 *
FROM Phase

DECLARE @comparisonDate AS DATE = CAST('2025-07-05 23:27:58.997' as DATE)
SELECT ph.Name, up.*
    FROM UserPhase up
    JOIN AspNetUsers a ON a.Id = up.UserId
    JOIN Phase ph ON ph.PhaseId = up.PhaseId
    WHERE DATEDIFF (DAY, up.ActivityDate, @comparisonDate) <= 7
        AND a.ProjectId = 46
        AND up.StatusCode LIKE '%9'
    ORDER BY up.ActivityDate DESC

USE arc;
SELECT *
FROM vDashboardDemographics	

SELECT q.GUID as 'UserID', q.Date as 'LastActivityDate', n.zip as 'PostalCode', CASE 
        WHEN (n.zip like ('[0-9][0-9][0-9][0-9][0-9]')) OR (n.zip like ('[0-9][0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]')) THEN LEFT(n.zip, 5)
		ELSE 'International'
    END AS 'PostalCodeState', n.state
FROM Dupuytrens.dbo.Dup80_pg11 q
JOIN arc.dbo.Natalog n ON q.guid = n.guid


SELECT countryname, max(isocountrycode)
FROM arc.dbo.Natalog
GROUP BY countryname
ORDER BY countryname
WHERE GUID = '0078DDC4-1316-43B1-850D-DA2E9D23B6CB'

SELECT primaddress, city, zip, state, country, countryname, isocountrycode
FROM arc.dbo.Natalog
WHERE zip like ('[0-9][0-9][0-9][0-9][0-9]')
	AND state NOT IN ('AK','AL','AR','AS','AZ','CA','CO','CT','DC','DE','FL','FM','GA','GU','HI','IA','ID','IL','IN','KS','KY','LA','MA','MD','ME','MH','MI','MN','MO','MP','MS','MT','NC','ND','NE','NH','NJ','NM','NV','NY','OH','OK','OR','PA','PR','PW','RI','SC','SD','TN','TX','UT','VA','VI','VT','WA','WI','WV','WY')
	AND country != 0
	AND isocountrycode NOT IN (840,630) 

SET @sql = N'
		WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', n.zip as ' + QUOTENAME('PostalCode') + '
						   FROM ' + @currentTableName + ' q
						   JOIN arc.dbo.Natalog n ON q.guid = n.guid) 

		INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode)
		SELECT UserID, LastActivityDate, SurveyStatus, PostalCode
		FROM pageQuery';


SET @sql = N'
		WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', CASE 
		                       WHEN n.state IN (''AK'',''AL'',''AR'',''AS'',''AZ'',''CA'',''CO'',''CT'',''DC'',''DE'',''FL'',''FM'',''GA'',''GU'',''HI'',''IA'',''ID'',''IL'',''IN'',''KS'',''KY'',''LA'',''MA'',''MD'',''ME'',''MH'',''MI'',''MN'',''MO'',''MP'',''MS'',''MT'',''NC'',''ND'',''NE'',''NH'',''NJ'',''NM'',''NV'',''NY'',''OH'',''OK'',''OR'',''PA'',''PR'',''PW'',''RI'',''SC'',''SD'',''TN'',''TX'',''UT'',''VA'',''VI'',''VT'',''WA'',''WI'',''WV'',''WY'') THEN n.state 
		                       WHEN (n.zip like (''[0-9][0-9][0-9][0-9][0-9]'')) OR (n.zip like (''[0-9][0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]'')) THEN LEFT(n.zip, 5) 
						       ELSE ''International'' 
						   END AS ' + QUOTENAME('PostalCode') + '
						   FROM ' + @currentTableName + ' q
						   JOIN arc.dbo.Natalog n ON q.guid = n.guid) 

		INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode)
		SELECT UserID, LastActivityDate, SurveyStatus, PostalCode
		FROM pageQuery';
USE Forward; DROP TABLE IF EXISTS #pageQueryResults CREATE TABLE #pageQueryResults (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT); DROP TABLE IF EXISTS #pageDataFull CREATE TABLE #pageDataFull (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT); DECLARE @currentTableName varchar(max) = ''; DECLARE @currentTableNameShort varchar(max) = ''; DECLARE @currentPageNum INT; DECLARE @currentPhaseStatus varchar(max) = ''; DECLARE @currentDatabasePrefix varchar(max) = ''; DECLARE @sql varchar(max) = ''; DECLARE tableListCursor CURSOR FOR SELECT DISTINCT pj.Name + '.dbo.' + ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END AS 'PageTable', ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END as 'TableName', ph.DatabasePrefix, p.PageNumber, nsr.ColumnTitle as 'PhaseStatus' FROM Forward.dbo.Phase ph JOIN Forward.dbo.Project pj ON pj.ProjectId = ph.ProjectId JOIN Forward.dbo.Page p ON p.SurveyToolId = ph.SurveyToolId JOIN Forward.dbo.Section s ON s.PageId = p.PageId JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId WHERE pj.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'Dupuytrens') AND pcc.Code = 'DUP' ORDER BY p.PageNumber; OPEN tableListCursor; FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; WHILE @@FETCH_STATUS = 0 BEGIN SET @sql = N' WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', n.zip as ' + QUOTENAME('PostalCode') + ', n.state as ' + QUOTENAME('StateAbbr') + ', n.country as ' + QUOTENAME('CountryCode') + ' FROM ' + @currentTableName + ' q JOIN arc.dbo.Natalog n ON q.guid = n.guid) INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode FROM pageQuery'; EXEC(@sql); WITH pageDataFullQuery AS (SELECT @currentDatabasePrefix AS 'DatabasePrefixVar', 'pg' + CASE WHEN @currentPageNum > 9 THEN CAST(@currentPageNum AS VARCHAR(3)) ELSE '0' + CAST(@currentPageNum AS VARCHAR(3)) END as 'PageNumber', pqr.UserID, pqr.LastActivityDate, pqr.SurveyStatus, pqr.PostalCode, pqr.StateAbbr, pqr.CountryCode FROM #pageQueryResults pqr) INSERT INTO #pageDataFull (DatabasePrefix, PageNum, UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT * FROM pageDataFullQuery DELETE FROM #pageQueryResults FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; END CLOSE tableListCursor; DEALLOCATE tableListCursor; DROP TABLE IF EXISTS #pageDataFinal CREATE TABLE #pageDataFinal (UserID VARCHAR(MAX), PostalCode VARCHAR(20), StateAbbr VARCHAR(20), CountryCode INT, DatabasePrefix VARCHAR(MAX), pg01_SurveyStatus INT, pg01_LastActivityDate DATETIME, pg02_SurveyStatus INT, pg02_LastActivityDate DATETIME, pg03_SurveyStatus INT, pg03_LastActivityDate DATETIME, pg04_SurveyStatus INT, pg04_LastActivityDate DATETIME, pg05_SurveyStatus INT, pg05_LastActivityDate DATETIME, pg06_SurveyStatus INT, pg06_LastActivityDate DATETIME, pg07_SurveyStatus INT, pg07_LastActivityDate DATETIME, pg08_SurveyStatus INT, pg08_LastActivityDate DATETIME, pg09_SurveyStatus INT, pg09_LastActivityDate DATETIME, pg10_SurveyStatus INT, pg10_LastActivityDate DATETIME, pg11_SurveyStatus INT, pg11_LastActivityDate DATETIME, pg12_SurveyStatus INT, pg12_LastActivityDate DATETIME, pg13_SurveyStatus INT, pg13_LastActivityDate DATETIME, pg14_SurveyStatus INT, pg14_LastActivityDate DATETIME, pg15_SurveyStatus INT, pg15_LastActivityDate DATETIME, pg16_SurveyStatus INT, pg16_LastActivityDate DATETIME, pg17_SurveyStatus INT, pg17_LastActivityDate DATETIME, pg18_SurveyStatus INT, pg18_LastActivityDate DATETIME, pg19_SurveyStatus INT, pg19_LastActivityDate DATETIME, pg20_SurveyStatus INT, pg20_LastActivityDate DATETIME, pg21_SurveyStatus INT, pg21_LastActivityDate DATETIME, pg22_SurveyStatus INT, pg22_LastActivityDate DATETIME, pg23_SurveyStatus INT, pg23_LastActivityDate DATETIME, pg24_SurveyStatus INT, pg24_LastActivityDate DATETIME); INSERT INTO #pageDataFinal SELECT pdfss.UserID, pdfss.PostalCode AS 'PostalCode', pdfss.StateAbbr AS 'StateAbbr', pdfss.CountryCode AS 'CountryCode', pdfss.DatabasePrefix, pdfss.pg01 AS 'pg01_SurveyStatus', pdflad.pg01 AS 'pg01_LastActivityDate', pdfss.pg02 AS 'pg02_SurveyStatus', pdflad.pg02 AS 'pg02_LastActivityDate', pdfss.pg03 AS 'pg03_SurveyStatus', pdflad.pg03 AS 'pg03_LastActivityDate', pdfss.pg04 AS 'pg04_SurveyStatus', pdflad.pg04 AS 'pg04_LastActivityDate', pdfss.pg05 AS 'pg05_SurveyStatus', pdflad.pg05 AS 'pg05_LastActivityDate', pdfss.pg06 AS 'pg06_SurveyStatus', pdflad.pg06 AS 'pg06_LastActivityDate', pdfss.pg07 AS 'pg07_SurveyStatus', pdflad.pg07 AS 'pg07_LastActivityDate', pdfss.pg08 AS 'pg08_SurveyStatus', pdflad.pg08 AS 'pg08_LastActivityDate', pdfss.pg09 AS 'pg09_SurveyStatus', pdflad.pg09 AS 'pg09_LastActivityDate', pdfss.pg10 AS 'pg10_SurveyStatus', pdflad.pg10 AS 'pg10_LastActivityDate', pdfss.pg11 AS 'pg11_SurveyStatus', pdflad.pg11 AS 'pg11_LastActivityDate', pdfss.pg12 AS 'pg12_SurveyStatus', pdflad.pg12 AS 'pg12_LastActivityDate', pdfss.pg13 AS 'pg13_SurveyStatus', pdflad.pg13 AS 'pg13_LastActivityDate', pdfss.pg14 AS 'pg14_SurveyStatus', pdflad.pg14 AS 'pg14_LastActivityDate', pdfss.pg15 AS 'pg15_SurveyStatus', pdflad.pg15 AS 'pg15_LastActivityDate', pdfss.pg16 AS 'pg16_SurveyStatus', pdflad.pg16 AS 'pg16_LastActivityDate', pdfss.pg17 AS 'pg17_SurveyStatus', pdflad.pg17 AS 'pg17_LastActivityDate', pdfss.pg18 AS 'pg18_SurveyStatus', pdflad.pg18 AS 'pg18_LastActivityDate', pdfss.pg19 AS 'pg19_SurveyStatus', pdflad.pg19 AS 'pg19_LastActivityDate', pdfss.pg20 AS 'pg20_SurveyStatus', pdflad.pg20 AS 'pg20_LastActivityDate', pdfss.pg21 AS 'pg21_SurveyStatus', pdflad.pg21 AS 'pg21_LastActivityDate', pdfss.pg22 AS 'pg22_SurveyStatus', pdflad.pg22 AS 'pg22_LastActivityDate', pdfss.pg23 AS 'pg23_SurveyStatus', pdflad.pg23 AS 'pg23_LastActivityDate', pdfss.pg24 AS 'pg24_SurveyStatus', pdflad.pg24 AS 'pg24_LastActivityDate' FROM (SELECT DatabasePrefix, UserID, MAX(PostalCode) AS 'PostalCode', MAX(StateAbbr) AS 'StateAbbr', MAX(CountryCode) AS 'CountryCode', MAX(pg01) AS 'pg01', MAX(pg02) AS 'pg02', MAX(pg03) AS 'pg03', MAX(pg04) AS 'pg04', MAX(pg05) AS 'pg05', MAX(pg06) AS 'pg06', MAX(pg07) AS 'pg07', MAX(pg08) AS 'pg08', MAX(pg09) AS 'pg09', MAX(pg10) AS 'pg10', MAX(pg11) AS 'pg11', MAX(pg12) AS 'pg12', MAX(pg13) AS 'pg13', MAX(pg14) AS 'pg14', MAX(pg15) AS 'pg15', MAX(pg16) AS 'pg16', MAX(pg17) AS 'pg17', MAX(pg18) AS 'pg18', MAX(pg19) AS 'pg19', MAX(pg20) AS 'pg20', MAX(pg21) AS 'pg21', MAX(pg22) AS 'pg22', MAX(pg23) AS 'pg23', MAX(pg24) AS 'pg24' FROM (SELECT * FROM #pageDataFull PIVOT (MAX(SurveyStatus) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdfss_pre GROUP BY DatabasePrefix, UserID) pdfss JOIN (SELECT * FROM #pageDataFull PIVOT (MAX(LastActivityDate) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdflad ON pdfss.UserID = pdflad.UserID AND pdfss.DatabasePrefix = pdflad.DatabasePrefix ORDER BY UserID, DatabasePrefix SELECT * FROM #pageDataFinal
USE Forward; DROP TABLE IF EXISTS #pageQueryResults CREATE TABLE #pageQueryResults (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(20), StateAbbr VARCHAR(20), CountryCode INT); DROP TABLE IF EXISTS #pageDataFull CREATE TABLE #pageDataFull (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(20), StateAbbr VARCHAR(20), CountryCode INT); DECLARE @currentTableName varchar(max) = ''; DECLARE @currentTableNameShort varchar(max) = ''; DECLARE @currentPageNum INT; DECLARE @currentPhaseStatus varchar(max) = ''; DECLARE @currentDatabasePrefix varchar(max) = ''; DECLARE @sql varchar(max) = ''; DECLARE tableListCursor CURSOR FOR SELECT DISTINCT pj.Name + '.dbo.' + ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END AS 'PageTable', ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END as 'TableName', ph.DatabasePrefix, p.PageNumber, nsr.ColumnTitle as 'PhaseStatus' FROM Forward.dbo.Phase ph JOIN Forward.dbo.Project pj ON pj.ProjectId = ph.ProjectId JOIN Forward.dbo.Page p ON p.SurveyToolId = ph.SurveyToolId JOIN Forward.dbo.Section s ON s.PageId = p.PageId JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId WHERE pj.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'Dupuytrens') AND pcc.Code = 'DUP' ORDER BY p.PageNumber; OPEN tableListCursor; FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; WHILE @@FETCH_STATUS = 0 BEGIN SET @sql = N' WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', n.zip as ' + QUOTENAME('PostalCode') + ', n.state as ' + QUOTENAME('StateAbbr') + ', n.country as ' + QUOTENAME('CountryCode') + ' FROM ' + @currentTableName + ' q JOIN arc.dbo.Natalog n ON q.guid = n.guid) INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode FROM pageQuery'; EXEC(@sql); WITH pageDataFullQuery AS (SELECT @currentDatabasePrefix AS 'DatabasePrefixVar', 'pg' + CASE WHEN @currentPageNum > 9 THEN CAST(@currentPageNum AS VARCHAR(3)) ELSE '0' + CAST(@currentPageNum AS VARCHAR(3)) END as 'PageNumber', pqr.UserID, pqr.LastActivityDate, pqr.SurveyStatus, pqr.PostalCode, pqr.StateAbbr, pqr.CountryCode FROM #pageQueryResults pqr) INSERT INTO #pageDataFull (DatabasePrefix, PageNum, UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT * FROM pageDataFullQuery DELETE FROM #pageQueryResults FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; END CLOSE tableListCursor; DEALLOCATE tableListCursor; 
USE Forward; DROP TABLE IF EXISTS #pageQueryResults CREATE TABLE #pageQueryResults (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(20), StateAbbr VARCHAR(20), CountryCode INT); DROP TABLE IF EXISTS #pageDataFull CREATE TABLE #pageDataFull (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(20), StateAbbr VARCHAR(20), CountryCode INT); DECLARE @currentTableName varchar(max) = ''; DECLARE @currentTableNameShort varchar(max) = ''; DECLARE @currentPageNum INT; DECLARE @currentPhaseStatus varchar(max) = ''; DECLARE @currentDatabasePrefix varchar(max) = ''; DECLARE @sql varchar(max) = ''; DECLARE tableListCursor CURSOR FOR SELECT DISTINCT pj.Name + '.dbo.' + ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END AS 'PageTable', ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END as 'TableName', ph.DatabasePrefix, p.PageNumber, nsr.ColumnTitle as 'PhaseStatus' FROM Forward.dbo.Phase ph JOIN Forward.dbo.Project pj ON pj.ProjectId = ph.ProjectId JOIN Forward.dbo.Page p ON p.SurveyToolId = ph.SurveyToolId JOIN Forward.dbo.Section s ON s.PageId = p.PageId JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId WHERE pj.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'Dupuytrens') AND pcc.Code = 'DUP' ORDER BY p.PageNumber; OPEN tableListCursor; FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; WHILE @@FETCH_STATUS = 0 BEGIN SET @sql = N' WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', n.zip as ' + QUOTENAME('PostalCode') + ', n.state as ' + QUOTENAME('StateAbbr') + ', n.country as ' + QUOTENAME('CountryCode') + ' FROM ' + @currentTableName + ' q JOIN arc.dbo.Natalog n ON q.guid = n.guid) INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode FROM pageQuery'; EXEC(@sql); WITH pageDataFullQuery AS (SELECT @currentDatabasePrefix AS 'DatabasePrefixVar', 'pg' + CASE WHEN @currentPageNum > 9 THEN CAST(@currentPageNum AS VARCHAR(3)) ELSE '0' + CAST(@currentPageNum AS VARCHAR(3)) END as 'PageNumber', pqr.UserID, pqr.LastActivityDate, pqr.SurveyStatus, pqr.PostalCode, pqr.StateAbbr, pqr.CountryCode FROM #pageQueryResults pqr) INSERT INTO #pageDataFull (DatabasePrefix, PageNum, UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT * FROM pageDataFullQuery DELETE FROM #pageQueryResults FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; END CLOSE tableListCursor; DEALLOCATE tableListCursor; 
USE Forward; DROP TABLE IF EXISTS #pageQueryResults CREATE TABLE #pageQueryResults (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT); DROP TABLE IF EXISTS #pageDataFull CREATE TABLE #pageDataFull (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT); DECLARE @currentTableName varchar(max) = ''; DECLARE @currentTableNameShort varchar(max) = ''; DECLARE @currentPageNum INT; DECLARE @currentPhaseStatus varchar(max) = ''; DECLARE @currentDatabasePrefix varchar(max) = ''; DECLARE @sql varchar(max) = ''; DECLARE tableListCursor CURSOR FOR SELECT DISTINCT pj.Name + '.dbo.' + ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END AS 'PageTable', ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END as 'TableName', ph.DatabasePrefix, p.PageNumber, nsr.ColumnTitle as 'PhaseStatus' FROM Forward.dbo.Phase ph JOIN Forward.dbo.Project pj ON pj.ProjectId = ph.ProjectId JOIN Forward.dbo.Page p ON p.SurveyToolId = ph.SurveyToolId JOIN Forward.dbo.Section s ON s.PageId = p.PageId JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId WHERE pj.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'Dupuytrens') AND pcc.Code = 'DUP' ORDER BY p.PageNumber; OPEN tableListCursor; FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; WHILE @@FETCH_STATUS = 0 BEGIN SET @sql = N' WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', n.zip as ' + QUOTENAME('PostalCode') + ', n.state as ' + QUOTENAME('StateAbbr') + ', n.country as ' + QUOTENAME('CountryCode') + ' FROM ' + @currentTableName + ' q JOIN arc.dbo.Natalog n ON q.guid = n.guid) INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode FROM pageQuery'; EXEC(@sql); WITH pageDataFullQuery AS (SELECT @currentDatabasePrefix AS 'DatabasePrefixVar', 'pg' + CASE WHEN @currentPageNum > 9 THEN CAST(@currentPageNum AS VARCHAR(3)) ELSE '0' + CAST(@currentPageNum AS VARCHAR(3)) END as 'PageNumber', pqr.UserID, pqr.LastActivityDate, pqr.SurveyStatus, pqr.PostalCode, pqr.StateAbbr, pqr.CountryCode FROM #pageQueryResults pqr) INSERT INTO #pageDataFull (DatabasePrefix, PageNum, UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT * FROM pageDataFullQuery DELETE FROM #pageQueryResults FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; END CLOSE tableListCursor; DEALLOCATE tableListCursor;
USE Forward; DROP TABLE IF EXISTS #pageQueryResults CREATE TABLE #pageQueryResults (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT); DROP TABLE IF EXISTS #pageDataFull CREATE TABLE #pageDataFull (DatabasePrefix VARCHAR(MAX), PageNum VARCHAR(10), UserID VARCHAR(MAX), LastActivityDate DATETIME, SurveyStatus INT, PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT); DECLARE @currentTableName varchar(max) = ''; DECLARE @currentTableNameShort varchar(max) = ''; DECLARE @currentPageNum INT; DECLARE @currentPhaseStatus varchar(max) = ''; DECLARE @currentDatabasePrefix varchar(max) = ''; DECLARE @sql varchar(max) = ''; DECLARE tableListCursor CURSOR FOR SELECT DISTINCT pj.Name + '.dbo.' + ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END AS 'PageTable', ph.DatabasePrefix + '_' + CASE WHEN s.SectionTypeId = 5 THEN 'pgDrug' ELSE 'pg' + CAST(p.PageNumber as VARCHAR(3)) END as 'TableName', ph.DatabasePrefix, p.PageNumber, nsr.ColumnTitle as 'PhaseStatus' FROM Forward.dbo.Phase ph JOIN Forward.dbo.Project pj ON pj.ProjectId = ph.ProjectId JOIN Forward.dbo.Page p ON p.SurveyToolId = ph.SurveyToolId JOIN Forward.dbo.Section s ON s.PageId = p.PageId JOIN ProjectCategories pc on pj.ProjectId = pc.ProjectId JOIN ProjectCategory pcc on pc.ProjectCategoryId = pcc.ProjectCategoryId JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId WHERE pj.ProjectId = (SELECT ProjectId FROM Forward.dbo.Project WHERE Name = 'Dupuytrens') AND pcc.Code = 'DUP' ORDER BY p.PageNumber; OPEN tableListCursor; FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; WHILE @@FETCH_STATUS = 0 BEGIN SET @sql = N' WITH pageQuery AS (SELECT q.GUID as ' + QUOTENAME('UserID') + ', q.Date as ' + QUOTENAME('LastActivityDate') + ', n.' + @currentPhaseStatus + ' as ' + QUOTENAME('SurveyStatus') + ', n.zip as ' + QUOTENAME('PostalCode') + ', n.state as ' + QUOTENAME('StateAbbr') + ', n.country as ' + QUOTENAME('CountryCode') + ' FROM ' + @currentTableName + ' q JOIN arc.dbo.Natalog n ON q.guid = n.guid) INSERT INTO #pageQueryResults (UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode FROM pageQuery'; EXEC(@sql); WITH pageDataFullQuery AS (SELECT @currentDatabasePrefix AS 'DatabasePrefixVar', 'pg' + CASE WHEN @currentPageNum > 9 THEN CAST(@currentPageNum AS VARCHAR(3)) ELSE '0' + CAST(@currentPageNum AS VARCHAR(3)) END as 'PageNumber', pqr.UserID, pqr.LastActivityDate, pqr.SurveyStatus, pqr.PostalCode, pqr.StateAbbr, pqr.CountryCode FROM #pageQueryResults pqr) INSERT INTO #pageDataFull (DatabasePrefix, PageNum, UserID, LastActivityDate, SurveyStatus, PostalCode, StateAbbr, CountryCode) SELECT * FROM pageDataFullQuery DELETE FROM #pageQueryResults FETCH NEXT FROM tableListCursor INTO @currentTableName, @currentTableNameShort, @currentDatabasePrefix, @currentPageNum, @currentPhaseStatus; END CLOSE tableListCursor; DEALLOCATE tableListCursor; DROP TABLE IF EXISTS #pageDataFinal CREATE TABLE #pageDataFinal (UserID VARCHAR(MAX), PostalCode VARCHAR(40), StateAbbr VARCHAR(40), CountryCode INT, DatabasePrefix VARCHAR(MAX), pg01_SurveyStatus INT, pg01_LastActivityDate DATETIME, pg02_SurveyStatus INT, pg02_LastActivityDate DATETIME, pg03_SurveyStatus INT, pg03_LastActivityDate DATETIME, pg04_SurveyStatus INT, pg04_LastActivityDate DATETIME, pg05_SurveyStatus INT, pg05_LastActivityDate DATETIME, pg06_SurveyStatus INT, pg06_LastActivityDate DATETIME, pg07_SurveyStatus INT, pg07_LastActivityDate DATETIME, pg08_SurveyStatus INT, pg08_LastActivityDate DATETIME, pg09_SurveyStatus INT, pg09_LastActivityDate DATETIME, pg10_SurveyStatus INT, pg10_LastActivityDate DATETIME, pg11_SurveyStatus INT, pg11_LastActivityDate DATETIME, pg12_SurveyStatus INT, pg12_LastActivityDate DATETIME, pg13_SurveyStatus INT, pg13_LastActivityDate DATETIME, pg14_SurveyStatus INT, pg14_LastActivityDate DATETIME, pg15_SurveyStatus INT, pg15_LastActivityDate DATETIME, pg16_SurveyStatus INT, pg16_LastActivityDate DATETIME, pg17_SurveyStatus INT, pg17_LastActivityDate DATETIME, pg18_SurveyStatus INT, pg18_LastActivityDate DATETIME, pg19_SurveyStatus INT, pg19_LastActivityDate DATETIME, pg20_SurveyStatus INT, pg20_LastActivityDate DATETIME, pg21_SurveyStatus INT, pg21_LastActivityDate DATETIME, pg22_SurveyStatus INT, pg22_LastActivityDate DATETIME, pg23_SurveyStatus INT, pg23_LastActivityDate DATETIME, pg24_SurveyStatus INT, pg24_LastActivityDate DATETIME); INSERT INTO #pageDataFinal SELECT pdfss.UserID, pdfss.PostalCode AS 'PostalCode', pdfss.StateAbbr AS 'StateAbbr', pdfss.CountryCode AS 'CountryCode', pdfss.DatabasePrefix, pdfss.pg01 AS 'pg01_SurveyStatus', pdflad.pg01 AS 'pg01_LastActivityDate', pdfss.pg02 AS 'pg02_SurveyStatus', pdflad.pg02 AS 'pg02_LastActivityDate', pdfss.pg03 AS 'pg03_SurveyStatus', pdflad.pg03 AS 'pg03_LastActivityDate', pdfss.pg04 AS 'pg04_SurveyStatus', pdflad.pg04 AS 'pg04_LastActivityDate', pdfss.pg05 AS 'pg05_SurveyStatus', pdflad.pg05 AS 'pg05_LastActivityDate', pdfss.pg06 AS 'pg06_SurveyStatus', pdflad.pg06 AS 'pg06_LastActivityDate', pdfss.pg07 AS 'pg07_SurveyStatus', pdflad.pg07 AS 'pg07_LastActivityDate', pdfss.pg08 AS 'pg08_SurveyStatus', pdflad.pg08 AS 'pg08_LastActivityDate', pdfss.pg09 AS 'pg09_SurveyStatus', pdflad.pg09 AS 'pg09_LastActivityDate', pdfss.pg10 AS 'pg10_SurveyStatus', pdflad.pg10 AS 'pg10_LastActivityDate', pdfss.pg11 AS 'pg11_SurveyStatus', pdflad.pg11 AS 'pg11_LastActivityDate', pdfss.pg12 AS 'pg12_SurveyStatus', pdflad.pg12 AS 'pg12_LastActivityDate', pdfss.pg13 AS 'pg13_SurveyStatus', pdflad.pg13 AS 'pg13_LastActivityDate', pdfss.pg14 AS 'pg14_SurveyStatus', pdflad.pg14 AS 'pg14_LastActivityDate', pdfss.pg15 AS 'pg15_SurveyStatus', pdflad.pg15 AS 'pg15_LastActivityDate', pdfss.pg16 AS 'pg16_SurveyStatus', pdflad.pg16 AS 'pg16_LastActivityDate', pdfss.pg17 AS 'pg17_SurveyStatus', pdflad.pg17 AS 'pg17_LastActivityDate', pdfss.pg18 AS 'pg18_SurveyStatus', pdflad.pg18 AS 'pg18_LastActivityDate', pdfss.pg19 AS 'pg19_SurveyStatus', pdflad.pg19 AS 'pg19_LastActivityDate', pdfss.pg20 AS 'pg20_SurveyStatus', pdflad.pg20 AS 'pg20_LastActivityDate', pdfss.pg21 AS 'pg21_SurveyStatus', pdflad.pg21 AS 'pg21_LastActivityDate', pdfss.pg22 AS 'pg22_SurveyStatus', pdflad.pg22 AS 'pg22_LastActivityDate', pdfss.pg23 AS 'pg23_SurveyStatus', pdflad.pg23 AS 'pg23_LastActivityDate', pdfss.pg24 AS 'pg24_SurveyStatus', pdflad.pg24 AS 'pg24_LastActivityDate' FROM (SELECT DatabasePrefix, UserID, MAX(PostalCode) AS 'PostalCode', MAX(StateAbbr) AS 'StateAbbr', MAX(CountryCode) AS 'CountryCode', MAX(pg01) AS 'pg01', MAX(pg02) AS 'pg02', MAX(pg03) AS 'pg03', MAX(pg04) AS 'pg04', MAX(pg05) AS 'pg05', MAX(pg06) AS 'pg06', MAX(pg07) AS 'pg07', MAX(pg08) AS 'pg08', MAX(pg09) AS 'pg09', MAX(pg10) AS 'pg10', MAX(pg11) AS 'pg11', MAX(pg12) AS 'pg12', MAX(pg13) AS 'pg13', MAX(pg14) AS 'pg14', MAX(pg15) AS 'pg15', MAX(pg16) AS 'pg16', MAX(pg17) AS 'pg17', MAX(pg18) AS 'pg18', MAX(pg19) AS 'pg19', MAX(pg20) AS 'pg20', MAX(pg21) AS 'pg21', MAX(pg22) AS 'pg22', MAX(pg23) AS 'pg23', MAX(pg24) AS 'pg24' FROM (SELECT * FROM #pageDataFull PIVOT (MAX(SurveyStatus) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdfss_pre GROUP BY DatabasePrefix, UserID) pdfss JOIN (SELECT * FROM #pageDataFull PIVOT (MAX(LastActivityDate) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdflad ON pdfss.UserID = pdflad.UserID AND pdfss.DatabasePrefix = pdflad.DatabasePrefix ORDER BY UserID, DatabasePrefix SELECT * FROM #pageDataFinal
SELECT * 
FROM #pageDataFull
ORDER BY UserID, PageNum
DROP TABLE IF EXISTS #pageDataFinal CREATE TABLE #pageDataFinal (UserID VARCHAR(MAX), PostalCode VARCHAR(20), StateAbbr VARCHAR(20), CountryCode INT, DatabasePrefix VARCHAR(MAX), pg01_SurveyStatus INT, pg01_LastActivityDate DATETIME, pg02_SurveyStatus INT, pg02_LastActivityDate DATETIME, pg03_SurveyStatus INT, pg03_LastActivityDate DATETIME, pg04_SurveyStatus INT, pg04_LastActivityDate DATETIME, pg05_SurveyStatus INT, pg05_LastActivityDate DATETIME, pg06_SurveyStatus INT, pg06_LastActivityDate DATETIME, pg07_SurveyStatus INT, pg07_LastActivityDate DATETIME, pg08_SurveyStatus INT, pg08_LastActivityDate DATETIME, pg09_SurveyStatus INT, pg09_LastActivityDate DATETIME, pg10_SurveyStatus INT, pg10_LastActivityDate DATETIME, pg11_SurveyStatus INT, pg11_LastActivityDate DATETIME, pg12_SurveyStatus INT, pg12_LastActivityDate DATETIME, pg13_SurveyStatus INT, pg13_LastActivityDate DATETIME, pg14_SurveyStatus INT, pg14_LastActivityDate DATETIME, pg15_SurveyStatus INT, pg15_LastActivityDate DATETIME, pg16_SurveyStatus INT, pg16_LastActivityDate DATETIME, pg17_SurveyStatus INT, pg17_LastActivityDate DATETIME, pg18_SurveyStatus INT, pg18_LastActivityDate DATETIME, pg19_SurveyStatus INT, pg19_LastActivityDate DATETIME, pg20_SurveyStatus INT, pg20_LastActivityDate DATETIME, pg21_SurveyStatus INT, pg21_LastActivityDate DATETIME, pg22_SurveyStatus INT, pg22_LastActivityDate DATETIME, pg23_SurveyStatus INT, pg23_LastActivityDate DATETIME, pg24_SurveyStatus INT, pg24_LastActivityDate DATETIME); INSERT INTO #pageDataFinal SELECT pdfss.UserID, pdfss.PostalCode AS 'PostalCode', pdfss.StateAbbr AS 'StateAbbr', pdfss.CountryCode AS 'CountryCode', pdfss.DatabasePrefix, pdfss.pg01 AS 'pg01_SurveyStatus', pdflad.pg01 AS 'pg01_LastActivityDate', pdfss.pg02 AS 'pg02_SurveyStatus', pdflad.pg02 AS 'pg02_LastActivityDate', pdfss.pg03 AS 'pg03_SurveyStatus', pdflad.pg03 AS 'pg03_LastActivityDate', pdfss.pg04 AS 'pg04_SurveyStatus', pdflad.pg04 AS 'pg04_LastActivityDate', pdfss.pg05 AS 'pg05_SurveyStatus', pdflad.pg05 AS 'pg05_LastActivityDate', pdfss.pg06 AS 'pg06_SurveyStatus', pdflad.pg06 AS 'pg06_LastActivityDate', pdfss.pg07 AS 'pg07_SurveyStatus', pdflad.pg07 AS 'pg07_LastActivityDate', pdfss.pg08 AS 'pg08_SurveyStatus', pdflad.pg08 AS 'pg08_LastActivityDate', pdfss.pg09 AS 'pg09_SurveyStatus', pdflad.pg09 AS 'pg09_LastActivityDate', pdfss.pg10 AS 'pg10_SurveyStatus', pdflad.pg10 AS 'pg10_LastActivityDate', pdfss.pg11 AS 'pg11_SurveyStatus', pdflad.pg11 AS 'pg11_LastActivityDate', pdfss.pg12 AS 'pg12_SurveyStatus', pdflad.pg12 AS 'pg12_LastActivityDate', pdfss.pg13 AS 'pg13_SurveyStatus', pdflad.pg13 AS 'pg13_LastActivityDate', pdfss.pg14 AS 'pg14_SurveyStatus', pdflad.pg14 AS 'pg14_LastActivityDate', pdfss.pg15 AS 'pg15_SurveyStatus', pdflad.pg15 AS 'pg15_LastActivityDate', pdfss.pg16 AS 'pg16_SurveyStatus', pdflad.pg16 AS 'pg16_LastActivityDate', pdfss.pg17 AS 'pg17_SurveyStatus', pdflad.pg17 AS 'pg17_LastActivityDate', pdfss.pg18 AS 'pg18_SurveyStatus', pdflad.pg18 AS 'pg18_LastActivityDate', pdfss.pg19 AS 'pg19_SurveyStatus', pdflad.pg19 AS 'pg19_LastActivityDate', pdfss.pg20 AS 'pg20_SurveyStatus', pdflad.pg20 AS 'pg20_LastActivityDate', pdfss.pg21 AS 'pg21_SurveyStatus', pdflad.pg21 AS 'pg21_LastActivityDate', pdfss.pg22 AS 'pg22_SurveyStatus', pdflad.pg22 AS 'pg22_LastActivityDate', pdfss.pg23 AS 'pg23_SurveyStatus', pdflad.pg23 AS 'pg23_LastActivityDate', pdfss.pg24 AS 'pg24_SurveyStatus', pdflad.pg24 AS 'pg24_LastActivityDate' FROM (SELECT DatabasePrefix, UserID, MAX(PostalCode) AS 'PostalCode', MAX(StateAbbr) AS 'StateAbbr', MAX(CountryCode) AS 'CountryCode', MAX(pg01) AS 'pg01', MAX(pg02) AS 'pg02', MAX(pg03) AS 'pg03', MAX(pg04) AS 'pg04', MAX(pg05) AS 'pg05', MAX(pg06) AS 'pg06', MAX(pg07) AS 'pg07', MAX(pg08) AS 'pg08', MAX(pg09) AS 'pg09', MAX(pg10) AS 'pg10', MAX(pg11) AS 'pg11', MAX(pg12) AS 'pg12', MAX(pg13) AS 'pg13', MAX(pg14) AS 'pg14', MAX(pg15) AS 'pg15', MAX(pg16) AS 'pg16', MAX(pg17) AS 'pg17', MAX(pg18) AS 'pg18', MAX(pg19) AS 'pg19', MAX(pg20) AS 'pg20', MAX(pg21) AS 'pg21', MAX(pg22) AS 'pg22', MAX(pg23) AS 'pg23', MAX(pg24) AS 'pg24' FROM (SELECT * FROM #pageDataFull PIVOT (MAX(SurveyStatus) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdfss_pre GROUP BY DatabasePrefix, UserID) pdfss JOIN (SELECT * FROM #pageDataFull PIVOT (MAX(LastActivityDate) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdflad ON pdfss.UserID = pdflad.UserID AND pdfss.DatabasePrefix = pdflad.DatabasePrefix ORDER BY UserID, DatabasePrefix SELECT * FROM #pageDataFinal
DROP TABLE IF EXISTS #pageDataFinal CREATE TABLE #pageDataFinal (UserID VARCHAR(MAX), PostalCode VARCHAR(20), StateAbbr VARCHAR(20), CountryCode INT, DatabasePrefix VARCHAR(MAX), pg01_SurveyStatus INT, pg01_LastActivityDate DATETIME, pg02_SurveyStatus INT, pg02_LastActivityDate DATETIME, pg03_SurveyStatus INT, pg03_LastActivityDate DATETIME, pg04_SurveyStatus INT, pg04_LastActivityDate DATETIME, pg05_SurveyStatus INT, pg05_LastActivityDate DATETIME, pg06_SurveyStatus INT, pg06_LastActivityDate DATETIME, pg07_SurveyStatus INT, pg07_LastActivityDate DATETIME, pg08_SurveyStatus INT, pg08_LastActivityDate DATETIME, pg09_SurveyStatus INT, pg09_LastActivityDate DATETIME, pg10_SurveyStatus INT, pg10_LastActivityDate DATETIME, pg11_SurveyStatus INT, pg11_LastActivityDate DATETIME, pg12_SurveyStatus INT, pg12_LastActivityDate DATETIME, pg13_SurveyStatus INT, pg13_LastActivityDate DATETIME, pg14_SurveyStatus INT, pg14_LastActivityDate DATETIME, pg15_SurveyStatus INT, pg15_LastActivityDate DATETIME, pg16_SurveyStatus INT, pg16_LastActivityDate DATETIME, pg17_SurveyStatus INT, pg17_LastActivityDate DATETIME, pg18_SurveyStatus INT, pg18_LastActivityDate DATETIME, pg19_SurveyStatus INT, pg19_LastActivityDate DATETIME, pg20_SurveyStatus INT, pg20_LastActivityDate DATETIME, pg21_SurveyStatus INT, pg21_LastActivityDate DATETIME, pg22_SurveyStatus INT, pg22_LastActivityDate DATETIME, pg23_SurveyStatus INT, pg23_LastActivityDate DATETIME, pg24_SurveyStatus INT, pg24_LastActivityDate DATETIME); INSERT INTO #pageDataFinal SELECT pdfss.UserID, pdfss.PostalCode AS 'PostalCode', pdfss.StateAbbr AS 'StateAbbr', pdfss.CountryCode AS 'CountryCode', pdfss.DatabasePrefix, pdfss.pg01 AS 'pg01_SurveyStatus', pdflad.pg01 AS 'pg01_LastActivityDate', pdfss.pg02 AS 'pg02_SurveyStatus', pdflad.pg02 AS 'pg02_LastActivityDate', pdfss.pg03 AS 'pg03_SurveyStatus', pdflad.pg03 AS 'pg03_LastActivityDate', pdfss.pg04 AS 'pg04_SurveyStatus', pdflad.pg04 AS 'pg04_LastActivityDate', pdfss.pg05 AS 'pg05_SurveyStatus', pdflad.pg05 AS 'pg05_LastActivityDate', pdfss.pg06 AS 'pg06_SurveyStatus', pdflad.pg06 AS 'pg06_LastActivityDate', pdfss.pg07 AS 'pg07_SurveyStatus', pdflad.pg07 AS 'pg07_LastActivityDate', pdfss.pg08 AS 'pg08_SurveyStatus', pdflad.pg08 AS 'pg08_LastActivityDate', pdfss.pg09 AS 'pg09_SurveyStatus', pdflad.pg09 AS 'pg09_LastActivityDate', pdfss.pg10 AS 'pg10_SurveyStatus', pdflad.pg10 AS 'pg10_LastActivityDate', pdfss.pg11 AS 'pg11_SurveyStatus', pdflad.pg11 AS 'pg11_LastActivityDate', pdfss.pg12 AS 'pg12_SurveyStatus', pdflad.pg12 AS 'pg12_LastActivityDate', pdfss.pg13 AS 'pg13_SurveyStatus', pdflad.pg13 AS 'pg13_LastActivityDate', pdfss.pg14 AS 'pg14_SurveyStatus', pdflad.pg14 AS 'pg14_LastActivityDate', pdfss.pg15 AS 'pg15_SurveyStatus', pdflad.pg15 AS 'pg15_LastActivityDate', pdfss.pg16 AS 'pg16_SurveyStatus', pdflad.pg16 AS 'pg16_LastActivityDate', pdfss.pg17 AS 'pg17_SurveyStatus', pdflad.pg17 AS 'pg17_LastActivityDate', pdfss.pg18 AS 'pg18_SurveyStatus', pdflad.pg18 AS 'pg18_LastActivityDate', pdfss.pg19 AS 'pg19_SurveyStatus', pdflad.pg19 AS 'pg19_LastActivityDate', pdfss.pg20 AS 'pg20_SurveyStatus', pdflad.pg20 AS 'pg20_LastActivityDate', pdfss.pg21 AS 'pg21_SurveyStatus', pdflad.pg21 AS 'pg21_LastActivityDate', pdfss.pg22 AS 'pg22_SurveyStatus', pdflad.pg22 AS 'pg22_LastActivityDate', pdfss.pg23 AS 'pg23_SurveyStatus', pdflad.pg23 AS 'pg23_LastActivityDate', pdfss.pg24 AS 'pg24_SurveyStatus', pdflad.pg24 AS 'pg24_LastActivityDate' FROM (SELECT DatabasePrefix, UserID, MAX(PostalCode) AS 'PostalCode', MAX(StateAbbr) AS 'StateAbbr', MAX(CountryCode) AS 'CountryCode', MAX(pg01) AS 'pg01', MAX(pg02) AS 'pg02', MAX(pg03) AS 'pg03', MAX(pg04) AS 'pg04', MAX(pg05) AS 'pg05', MAX(pg06) AS 'pg06', MAX(pg07) AS 'pg07', MAX(pg08) AS 'pg08', MAX(pg09) AS 'pg09', MAX(pg10) AS 'pg10', MAX(pg11) AS 'pg11', MAX(pg12) AS 'pg12', MAX(pg13) AS 'pg13', MAX(pg14) AS 'pg14', MAX(pg15) AS 'pg15', MAX(pg16) AS 'pg16', MAX(pg17) AS 'pg17', MAX(pg18) AS 'pg18', MAX(pg19) AS 'pg19', MAX(pg20) AS 'pg20', MAX(pg21) AS 'pg21', MAX(pg22) AS 'pg22', MAX(pg23) AS 'pg23', MAX(pg24) AS 'pg24' FROM (SELECT * FROM #pageDataFull PIVOT (MAX(SurveyStatus) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdfss_pre GROUP BY DatabasePrefix, UserID) pdfss JOIN (SELECT * FROM #pageDataFull PIVOT (MAX(LastActivityDate) FOR PageNum IN (pg01,pg02,pg03,pg04,pg05,pg06,pg07,pg08,pg09,pg10,pg11,pg12,pg13,pg14,pg15,pg16,pg17,pg18,pg19,pg20,pg21,pg22,pg23,pg24)) AS pageDataFullPivot) pdflad ON pdfss.UserID = pdflad.UserID AND pdfss.DatabasePrefix = pdflad.DatabasePrefix ORDER BY UserID, DatabasePrefix SELECT * FROM #pageDataFinal



