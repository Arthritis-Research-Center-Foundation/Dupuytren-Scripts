/*USE Forward
SELECT ph.Name as 'Phase Name', n.Diagnosis, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID
    JOIN Status s on s.Code = up.StatusCode
    JOIN Phase ph ON ph.PhaseId = up.PhaseId
    WHERE diagnosis = 'DUP'
        AND ph.Name = 'Dupuytrens Enrollment'
        AND s.Display = 'Complete'
        AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7
    GROUP BY ph.name, diagnosis*/

--USE Forward
----SELECT ph.Name as 'Phase Name', n.Diagnosis, COUNT(*) as Count
--SELECT ph.Name as 'Phase Name', n.Diagnosis, s.Display, s.code, s.IsComplete, ph.CompleteStatus, ph.PhaseID
--    FROM arc.dbo.Natalog n
--    JOIN UserPhase up ON up.UserId = n.GUID
--    JOIN Status s on s.Code = up.StatusCode
--    JOIN Phase ph ON ph.PhaseId = up.PhaseId
--    WHERE diagnosis = 'DUP'
--        --AND ph.Name = 'Dupuytrens Enrollment'
--        --AND s.Display = 'Complete'
--        --AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7
--    --GROUP BY ph.name, diagnosis
--	ORDER BY ph.CompleteStatus, ph.PhaseId DESC

-- Get list of unique Forward Phase.'phase names' and Phase.'IDs'
--USE Forward
--SELECT PhaseID as 'Phase ID', Name as 'Phase Name', Phase.StartedStatus --, COUNT(*) AS [Participants in Phase]
--FROM Phase
----WHERE Phase.Name Like 
----	CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July...
----		CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'January'
----	ELSE
----		CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'July'
----	END
----GROUP BY PhaseID, Name
--ORDER BY Name

-- Get list of unique Forward UserPhase.UserIDs
--USE Forward
--SELECT DISTINCT PhaseID
--FROM UserPhase
--ORDER BY PhaseID

-- Get list of unique Forward Status.Codes
/**/

----------------------------------------------------------------------------------------------------------------------------
-- Enrollment: Total Started this Week - This is the total number of people who have started a Dupuytren Enrollment questionnaire this week --
USE Forward
SELECT ph.Name as 'Phase Name', n.Diagnosis, COUNT(*) as Count
SELECT COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name = 'Dupuytrens Enrollment' -- must be enrolled in Dupuytrens --
        AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --
    GROUP BY ph.name, diagnosis

-- Enrollment: Total Completed this Week - Same as the above, but for a Completed status --
USE Forward
SELECT ph.Name as 'Phase Name', n.Diagnosis, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name = 'Dupuytrens Enrollment' -- must be enrolled in Dupuytrens --
		AND s.IsComplete = 1 -- must have a status of complete --
        AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --
    GROUP BY ph.name, diagnosis

-- Enrollment: Total Started Ever - We are looking for the total number of people who have a Dupuytren Diagnosis and have ever started a Dupuytren Enrollment --
USE Forward
SELECT ph.Name as 'Phase Name', n.Diagnosis, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name = 'Dupuytrens Enrollment' -- must be enrolled in Dupuytrens --
    GROUP BY ph.name, diagnosis

-- Enrollment: Total Completed Ever - Same as the above, but for a Completed status --
USE Forward
SELECT ph.Name as 'Phase Name', n.Diagnosis, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name = 'Dupuytrens Enrollment' -- must be enrolled in Dupuytrens --
		AND s.IsComplete = 1 -- must have a status of complete --
    GROUP BY ph.name, diagnosis

-- Long Questionnaire: Current Phase Questionnaires Sent - This is the total number of phase questionnaires that have been sent out to Dupuytren participants this phase --
USE Forward
SELECT ph.Name as 'Phase Name', ph.SentStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.SentStatus -- confirm that the sent status is 'sent' --
	GROUP BY ph.Name, ph.SentStatus

-- Long Questionnaire: Current Phase Questionnaires Started - This is the total number of phase questionnaires that have been started by Dupuytren participants this phase --
USE Forward
SELECT ph.Name as 'Phase Name', ph.StartedStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.StartedStatus -- confirm that the started status is 'started' --
	GROUP BY ph.Name, ph.StartedStatus

-- Long Questionnaire: Current Phase Questionnaires Completed - This is the total number of phase questionnaires that have been completed by Dupuytren participants this phase --
USE Forward
SELECT ph.Name as 'Phase Name', ph.CompleteStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is 'completed' --
	GROUP BY ph.Name, ph.CompleteStatus

-- Long Questionnaire: Phase Questionnaires Completed this week - This is the total number of phase questionnaires that have been completed by Dupuytren participants this week --
USE Forward
SELECT ph.Name as 'Phase Name', ph.CompleteStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is 'complete' --
		AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --
	GROUP BY ph.Name, ph.CompleteStatus

-- Long Questionnaire: Total Phase Questionnaires Ever Sent - This is the total number of phase questionnaires that have ever been sent out to Dupuytren participants --
USE Forward
SELECT COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like '%Questionnaire (%Dupuytren%' -- specify the long questionnaire phase name... --
		AND ph.Name Not Like '%Questionnaire (%Dupuytren%Short%'
		AND up.StatusCode = ph.SentStatus -- confirm that the sent status is 'sent' --

-- Long Questionnaire: Total Phase Questionnaires Ever Completed - This is the total number of phase questionnaires that have ever been completed by Dupuytren participants --
USE Forward
SELECT COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like '%Questionnaire (%Dupuytren%' -- specify the long questionnaire phase name... --
		AND ph.Name Not Like '%Questionnaire (%Dupuytren%Short%'
		AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is 'complete' --

-- Short Questionnaire: Current Phase Questionnaires Sent - This is the total number of phase questionnaires that have been sent out to Dupuytren participants this phase --
USE Forward
SELECT ph.Name as 'Phase Name', ph.SentStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.SentStatus -- confirm that the sent status is 'sent' --
	GROUP BY ph.Name, ph.SentStatus

-- Short Questionnaire: Current Phase Questionnaires Started - This is the total number of phase questionnaires that have been started by Dupuytren participants this phase --
USE Forward
SELECT ph.Name as 'Phase Name', ph.StartedStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.StartedStatus -- confirm that the started status is 'started' --
	GROUP BY ph.Name, ph.StartedStatus

-- Short Questionnaire: Current Phase Questionnaires Completed - This is the total number of phase questionnaires that have been completed by Dupuytren participants this phase --
USE Forward
SELECT ph.Name as 'Phase Name', ph.CompleteStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE())-2, ' January Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE())-2, ' July Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is 'completed' --
	GROUP BY ph.Name, ph.CompleteStatus

-- Short Questionnaire: Phase Questionnaires Completed this week - This is the total number of phase questionnaires that have been completed by Dupuytren participants this week --
USE Forward
SELECT ph.Name as 'Phase Name', ph.CompleteStatus, COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
        AND ph.Name Like -- specify the current phase name... --
			CASE WHEN MONTH(GETDATE()) < 7 THEN -- if the current month is pre-July... --
				CONCAT(YEAR(GETDATE()), ' January Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'January' --
			ELSE
				CONCAT(YEAR(GETDATE()), ' July Questionnaire (Dupuytren Short)') -- search for a phase name that includes the current year and 'July' --
			END
		AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is 'complete' --
		AND DATEDIFF(day, up.ActivityDate, GETDATE()) <= 7 -- must have occured in the last week --
	GROUP BY ph.Name, ph.CompleteStatus

-- Short Questionnaire: Total Phase Questionnaires Ever Sent - This is the total number of phase questionnaires that have ever been sent out to Dupuytren participants --
USE Forward
SELECT COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
		AND ph.Name Like '%Questionnaire (Dupuytren Short)'
		AND up.StatusCode = ph.SentStatus -- confirm that the sent status is 'sent' --

-- Short Questionnaire: Total Phase Questionnaires Ever Completed - This is the total number of phase questionnaires that have ever been completed by Dupuytren participants --
USE Forward
SELECT COUNT(*) as Count
    FROM arc.dbo.Natalog n
    JOIN UserPhase up ON up.UserId = n.GUID -- match userIDs between Forward.UserPhase and arc.Natalog --
    JOIN Status s on s.Code = up.StatusCode -- match status codes between Forward.Status and Forward.UserPhase --
    JOIN Phase ph ON ph.PhaseId = up.PhaseId -- match phase IDs between Forward.Phase and Forward.UserPhase --
    WHERE diagnosis = 'DUP' -- must have DUP diagnosis --
		AND ph.Name Like '%Questionnaire (%Dupuytren Short)' -- specify the short questionnaire phase name... --
        AND up.StatusCode = ph.CompleteStatus -- confirm that the complete status is 'complete' --




