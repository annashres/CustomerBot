"use strict";
const builder = require('botbuilder');
const azure_builder = require('botbuilder-azure');

const connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Create chat bot
var bot = new builder.UniversalBot(connector);

//Add first run dialog
bot.dialog('/firstRun',
[
    function (session)
    {
        const botName = 'MAAV-1';
        const description = `I record notes and insights from the conversations Microsoft employees have with customers.  My purpose is to make it easier to disseminate knowledge about Microsoft customers and win as a team.`;

        session.send(`Hi there! I'm ${botName}`);
        session.send(`In a nutshell, here's what I do:\n\n${description}`);

        builder.Prompts.text(session, `What's your name?`);
    },
    function (session, results)
    {
        session.userData.name = results.response
        session.endDialog("Nice to meet you %s. I'm now going to ask you questions to record your customer conversation ...", session.userData.name); 
        session.beginDialog('/dataEntry');
    }
]);

// Main dialog loop
bot.dialog('/', [
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
bot.dialog('/dataEntry',
[
    function (session)
    {
        builder.Prompts.text(session, "Besides yourself, who else was on the call");
    },
    function (session, results)
    {
        session.userData.authors = results.response;
        builder.Prompts.text(session, "What company did y'all speak with?");
    },
    function (session, results)
    {
        session.userData.company = results.response;
        builder.Prompts.text(session, "Who did you speak with at " + session.userData.company + "?");
    },
    function (session, results)
    {
        session.userData.contact = results.response;
        builder.Prompts.choice(session, "Which products did you talk about?", ["SQLVM", "Elastic Pools", "Azure SQL DB", "On-Prem SQL Server", "Other"]);
    },
    function (session, results)
    {
        session.userData.product = results.response.entity;
        builder.Prompts.text(session, "Please paste in below any notes you took down during the call");
    },
    function (session, results)
    {
        session.userData.notes = results.response;
        builder.Prompts.text(session, "Any tags you would like to add to make this conversation easier to find in the future?");
    },
    function (session, results)
    {
        session.userData.tags = results.response;
        var conversationNotes = JSON.stringify(session.userData);
        session.send("That's all I need. Thanks for sending this info:");
        session.send("%s", conversationNotes);
        session.endDialog();
    }
]);

module.exports = { 
    connector: connector,
    bot: bot
};