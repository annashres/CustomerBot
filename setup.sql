/*
CREATE DATABASE TodoDb;
USE TodoDb;
*/
DROP TABLE IF EXISTS Feedback
DROP PROCEDURE IF EXISTS createFeedback
GO

CREATE TABLE Feedback (
	id int IDENTITY PRIMARY KEY,
	name nvarchar(256) NULL,
	authors nvarchar(256) NULL,
	company nvarchar(256) NULL,
	contact nvarchar(256) NULL,
	product nvarchar(256) NULL,
	notes nvarchar(4000) NULL,
	tags nvarchar(4000) NULL,

)
GO


create procedure dbo.createFeedback(@feedback nvarchar(max))
as begin
    INSERT INTO Feedback 
    SELECT *
    FROM OPENJSON(@feedback)
    with (name nvarchar(256), authors nvarchar(256), company nvarchar(256), contact nvarchar(256), 
    product nvarchar(256), notes nvarchar(4000), tags nvarchar(4000))
end
GO