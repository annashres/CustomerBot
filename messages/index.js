// Imports
"use strict";
var botframework = require('botbuilder');
var azure_botframework = require("botbuilder-azure");
var useEmulator = (process.env.NODE_ENV == 'development');

// Create default universal chat connector
var chatEnvironment = useEmulator ? new botframework.ChatConnector() : new azure_botframework.BotServiceConnector({
	appId: process.env.MicrosoftAppId,
	appPassword: process.env.MicrosoftAppPassword
});

// Create chat bot
var profilebot = new azure_botframework.UniversalBot(chatEnvironment);

// Add first run dialog
profilebot.dialog('firstRun', 
[
    function (session)
    {
        botframework.Prompts.text(session, "Aloha!! I'm MAAV-1, the Feedback bot... What's your name?");
    },
    function (session, results)
    {
        session.userData.name = results.response;
        session.endDialog("Nice to meet you %s. I'm now going to ask you questions to record your customer conversation ...", session.userData.name); 
        session.beginDialog('/dataEntry');
    }
]);

// Main dialog loop
profilebot.dialog('/',
[
	function (session, args, next)
	{
		if (!session.userData.name)
		{
			session.beginDialog('/firstRun');
		}
		else
		{
			session.send("Hello there %s! Nice to see you again. I'm going to ask you questions now to record your customer conversation ...", session.userData.name);
			session.beginDialog('/dataEntry');
		}
	}
]);

// Data entry dialog
profilebot.dialog('/dataEntry',
[
	function (session)
	{
		azure_botframework.Prompts.text(session, "Besides yourself, who else was on the call");
	},
	function (session, results)
	{
		session.userData.authors = results.response;
		azure_botframework.Prompts.text(session, "What company did y'all speak with?");
	},
	function (session, results)
	{
		session.userData.company = results.response;
		azure_botframework.Prompts.text(session, "Who did you speak with at %s?", session.userData.company);
	},
	function (session, results)
	{
		session.userData.contact = results.response;
		azure_botframework.Prompts.choice(session, "Which products did you talk about?", ["SQLVM", "Elastic Pools", "Azure SQL DB", "On-Prem SQL Server", "Other"]);
	},
	function (session, results)
	{
		session.userData.product = results.response.entity;
		azure_botframework.Prompts.text(session, "Please paste in below any notes you took down during the call");
	},
	function (session, results)
	{
		session.userData.notes = results.response;
		azure_botframework.Prompts.text(session, "Any tags you would like to add to make this conversation easier to find in the future?");
	},
	function (session, results)
	{
		session.userData.tags = results.response;
		var conversationNotes = JSON.stringify(userData);
		session.send("That's all I need. Thanks for sending this info:");
		session.send("%s", conversationNotes);
		session.endDialog();
	}
]);

// --- Local testing on dev machine ------
if (useEmulator)
{
	var restify = require('restify');
	var server = restify.createServer();
	server.listen(process.env.PORT || 3978, function()
	{
		console.log(' test bot endpoint listening at %s', server.url);
	});
	server.post('/api/messages', chatEnvironment.listen());
}
else
{
    module.exports = { default: chatEnvironment.listen() }
}
