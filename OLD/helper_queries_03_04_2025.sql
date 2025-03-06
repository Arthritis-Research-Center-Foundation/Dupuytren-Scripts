--USE Forward
--SELECT nsr.ColumnTitle as 'NatalogField', DatabasePrefix
--FROM Phase ph
--JOIN NatalogSurveyReference nsr on nsr.NatalogSurveyReferenceId = ph.NatalogSurveyReferenceId
--    AND ph.ProjectId = 46 -- 57 = PsO
--WHERE ph.DatabasePrefix like 'Dup%' -- must have {SURVEYNAME} database prefix --' ...



--USE Forward 
--SELECT distinct DatabasePrefix, '['' + CAST(PageNumber as VARCHAR(3)) + ''] ' + Pageheader as 'Page Header' 
--    FROM Page p 
--    JOIN SurveyTool st ON st.SurveyToolId = p.SurveyToolId -- match SurveyToolIDs between Forward.SurveyTools and Forward.Page --' ...
--    JOIN Phase ph ON ph.PhaseId = st.PhaseId -- match phase IDs between Forward.Phase and Forward.SurveyTools --' ...
--    WHERE ph.DatabasePrefix like 'Dup%' -- must have {SURVEYNAME} database prefix --' ...
--		AND ProjectId = 46 -- specify the projectID... --




--USE Dupuytrens
--SELECT TABLE_NAME
--FROM Dupuytrens.INFORMATION_SCHEMA.TABLES
--WHERE TABLE_NAME like 'Dup%_pg%' -- must have {SURVEYNAME} in table name --


--USE Forward 
--SELECT COUNT(*) AS 'Questionnaires sent out'
--FROM UserPhase up -- grab user activity --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
--	AND {SURVEYSELECT} -- limit UserPhase to only include the {SURVEYDESCRIPTION} --


USE Forward 
SELECT COUNT(*) AS 'Num Participants that Completed Page'
FROM Dupuytrens.dbo.Dup87_pg1 q -- grab info for specific page of specific questionnaire --' ...
JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...
	AND ph.DatabasePrefix = 'Dup87' -- limit UserPhase to only include the {SURVEYDESCRIPTION} --', ...
	--AND Complete = 1 -- only include complete pages --
	--AND jul24 = ph.CompleteStatus

--USE Forward 
--SELECT COUNT(*) AS 'Num Participants that Completed Page' 
--FROM Dupuytrens.dbo.Dup87_pgDrug q -- grab info for specific page of specific questionnaire --
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.Dup87_pgDrug and Forward.UserPhase --
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --
--	AND ph.DatabasePrefix = 'Dup87' -- limit UserPhase to only include the Dupuytrens Long Survey for ph87 --
--	AND Complete = 1 -- only include complete pages --    
--	AND jul24 = ph.CompleteStatus









--USE Forward 
--SELECT COUNT(*) AS 'Num Participants that Completed Page'
--FROM Dupuytrens.dbo.Dup87_pg1 q -- grab info for specific page of specific questionnaire --' ...
--JOIN UserPhase up ON q.guid = up.UserId -- match userIDs between Dupuytrens.dbo.{SURVEYPAGEHEADER} and Forward.UserPhase --' ...
--JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --' ...
--JOIN Project pj ON pj.ProjectId = ph.ProjectId -- match projectIDs between Forward.Project and Forward.Phase --' ...
--JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database --' ...
--WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity --' ...
--	AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project) --' ...









