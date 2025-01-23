use Forward;

----------------------------------------
-- Dupuytren Enrollment Questionnaire --
----------------------------------------

-- get the number of people who should have completed questionnaires
SELECT COUNT(*) AS 'COUNT - Enrollments sent out'
	FROM UserPhase up -- grab user activity --
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.Name LIKE '%Enrollment%' -- limit UserPhase to only include the Dupuytrens Enrollment PhaseId
		
-- get the number of people per page who started the page
SELECT COUNT(*) AS 'COUNT - Started enrollment page one' 
	FROM Dupuytrens.dbo.EnrollDup_pg1 e
	JOIN UserPhase up ON e.guid = up.UserId
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.Name LIKE '%Enrollment%' -- limit UserPhase to only include the Dupuytrens Enrollment PhaseId


-- get the number of people who completed the page
SELECT COUNT(*) AS 'COUNT - Completed enrollment page one' 
	FROM Dupuytrens.dbo.EnrollDup_pg1 e
	JOIN UserPhase up ON e.guid = up.UserId
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.Name LIKE '%Enrollment%' -- limit UserPhase to only include the Dupuytrens Enrollment PhaseId
		AND Complete = 1

-------------------------------------------
-- Dupuytren Long Questionnaire for ph87 --
-------------------------------------------

-- get the number of people who should have completed questionnaires
SELECT COUNT(*) AS 'COUNT - Dupuytren Long sent out for ph87'
	FROM UserPhase up -- grab user activity --
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.DatabasePrefix = 'Dup87' -- limit UserPhase to only include the Dupuytrens Long Survey for ph87
		
-- get the number of people per page who started the page
SELECT COUNT(*) AS 'COUNT - Started ph87 Long page one' 
	FROM Dupuytrens.dbo.Dup87_pg1 q
	JOIN UserPhase up ON q.guid = up.UserId
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.DatabasePrefix = 'Dup87' -- limit UserPhase to only include the Dupuytrens Short Survey for ph87


-- get the number of people who completed the page
SELECT COUNT(*) AS 'COUNT - Completed ph87 Long page one' 
	FROM Dupuytrens.dbo.Dup87_pg1 q
	JOIN UserPhase up ON q.guid = up.UserId
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.DatabasePrefix = 'Dup87' -- limit UserPhase to only include the Dupuytrens Short Survey for ph87
		AND Complete = 1

--------------------------------------------
-- Dupuytren Short Questionnaire for ph87 --
--------------------------------------------

-- get the number of people who should have completed questionnaires
SELECT COUNT(*) AS 'COUNT - Dupuytren Short sent out for ph87'
	FROM UserPhase up -- grab user activity --
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.DatabasePrefix = 'Short87' -- limit UserPhase to only include the Dupuytrens Short Survey for ph87
		
-- get the number of people per page who started the page
SELECT COUNT(*) AS 'COUNT - Started ph87 Short page one' 
	FROM Dupuytrens.dbo.Short87_pg1 q
	JOIN UserPhase up ON q.guid = up.UserId
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.DatabasePrefix = 'Short87' -- limit UserPhase to only include the Dupuytrens Short Survey for ph87

-- get the number of people who completed the page
SELECT COUNT(*) AS 'COUNT - Completed ph87 Short page one' 
	FROM Dupuytrens.dbo.Short87_pg1 q
	JOIN UserPhase up ON q.guid = up.UserId
	JOIN Phase ph ON ph.PhaseId = up.PhaseId 
	JOIN Project pj ON pj.ProjectId = ph.ProjectId
	JOIN arc.dbo.Natalog n ON n.guid = up.UserId -- joining on Natalog ensures we only include participants still in the database
	WHERE up.StatusCode = ph.SentStatus -- only keep "Sent" activity
		AND pj.Name = 'Dupuytrens' -- limit UserPhase to only include the Dupuytrens Registry (aka Project)
		AND ph.DatabasePrefix = 'Short87' -- limit UserPhase to only include the Dupuytrens Short Survey for ph87
		AND Complete = 1
