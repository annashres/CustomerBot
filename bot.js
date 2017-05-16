"use strict";
var builder = require('botbuilder');
var azure_builder = require('botbuilder-azure');
var Sequelize = require('sequelize');

var userName = 'sa';
var password = PASSWORD; 
var hostName = 'localhost';
var sampleDbName = DATABASE;

var sampleDb = new Sequelize(sampleDbName, userName, password, {
    dialect: 'mssql',
    host: hostName,
    port: 1433, // Default port
    logging: false, // disable logging; default: console.log

    dialectOptions: {
        requestTimeout: 30000 // timeout = 30 seconds
    }
});

// Define the 'Feedback' model
var Feedback = sampleDb.define('feedback', {
    Name: Sequelize.STRING,
    Authors: Sequelize.STRING,
    Company: Sequelize.STRING,
    Contact: Sequelize.STRING,
    Product: Sequelize.STRING,
    Notes: Sequelize.STRING,
    Tags: Sequelize.STRING
});

sampleDb.sync()
.then(function() {
    console.log('\nCreated database schema from model.');
})
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Create chat bot
var bot = new builder.UniversalBot(connector);

// Welcome message
bot.on('conversationUpdate', function(message)
{
    if (message.membersAdded)
    {
        message.membersAdded.forEach(function(identity)
        {
            if (identity.id === message.address.bot.id)
            {
                console.log('Sending greeting message...')
                bot.beginDialog(message.address, '/');
            }
        });
    }
});

//Add first run dialog
bot.dialog('/firstRun',
[
    function (session)
    {
        var botName = 'MAAV-1';
        var description = `I record notes and insights from the conversations Microsoft employees have with customers.  My purpose is to make it easier to disseminate knowledge about Microsoft customers and win as a team.`;

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
            session.send("Welcome back %s! I'm guessing you have a new customer conversation for me. I'm going to ask you questions now to record your customer conversation ...", session.userData.name);
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
        session.send("That's all I need.");
        
        session.beginDialog('/conversationCard');
    } 
]);

// Confirm data entry dialog
bot.dialog('/conversationCard', [
    function(session, args, next)
    {
        // Show the details that were entered if available
        if (!session.message.value)
        {
            session.send("Please confirm the information below is accurate");
                
            var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + session.userData.contact + " today where you discussed about how " + session.userData.product + " is used at " + session.userData.company + "</s><voice gender = \"female\"></voice>"
            var header = "Conversation with " + session.userData.company
            var authors = session.userData.name + ", " + session.userData.authors

            var outputCard = new builder.Message(session)
            .addAttachment({
                contentType: "application/vnd.microsoft.card.adaptive",
                content:
                {
                    type: "AdaptiveCard",
                    speak: audioSummary,
                    body:
                    [
                        {
                            "type": "TextBlock",
                            "text": header,
                            "size": "large",
                            "weight": "bolder"
                        },
                        {
                            "type": "FactSet",
                            "facts":
                            [
                                {
                                    "title": "Company:",
                                    "value": session.userData.company
                                },
                                {
                                    "title": "Author(s):",
                                    "value": authors
                                },
                                {
                                    "title": "Customer contact(s):",
                                    "value": session.userData.contact
                                },
                                {
                                    "title": "Product(s) discussed:",
                                    "value": session.userData.product
                                },
                                {
                                    "title": "Tags:",
                                    "value": session.userData.tags
                                },
                                {
                                    "title": "Notes:",
                                    "value": session.userData.notes
                                }

                            ]
                        }                 
                    ],
                    actions:
                    [
                        {
                            "type": "Action.ShowCard",
                            "title": "Edit conversation",
                            "card":
                            {
                                "type": "AdaptiveCard",
                                "body":
                                [
                                    {
                                        "type": "Input.Text",
                                        "id": "comment",
                                        "isMultiline": true,
                                        "placeholder": "Enter your change here. E.g. Product: MySql DB"
                                    }
                                ],
                                "actions": 
                                [
                                    {
                                        "type": "Action.Submit",
                                        "title": "OK",
                                        "data": {"message": "edit"}
                                    }
                                ]
                            }
                        },
                        {
                            "type": "Action.Submit",
                            "title": "Confirm",
                            "data": { "message" : "confirm"}  
                        }
                    ]
                }
            });

            session.send(outputCard);
        }
        
        if ((session.message) && (session.message.value))
        {
            var response = session.message.value;
            
            if (response.message === "edit")
                session.replaceDialog('/editConversation');
            else if (response.message === "confirm"){
                session.replaceDialog('/confirm');
                Feedback.create({Name: session.userData.name, Authors: session.userData.authors, Company: session.userData.company, Contact: session.userData.contact, Product: session.userData.product, Notes: session.userData.notes, Tags: session.userData.tags}).then(feedback => {
                console.log(feedback.get({
                    plain: true
                }))
                });
            }
            else
            {
                session.message.value = null;
                session.send("Did not understand that response");
            }    
        }

    }
]);

// Edit existing conversation details
bot.dialog('/editConversation', [
    function (session, args, next)
    {
        var inputMessage = session.message.value.comment

        inputMessage = inputMessage.split(':');
        var inputTag = inputMessage[0];
        var inputValue = inputMessage[1];

        
        session.send("Updating '%s' entry", inputTag);
        session.userData[inputTag] = inputValue;
        session.message.value = null;
        console.log(session.userData);

        session.beginDialog('/conversationCard');
    }
]);

// End dialog
bot.dialog('/confirm', [
    function (session, args, next)
    {
        session.send("Goodbye now")
        session.endDialog();
    }
]);

//Help dialog
bot.dialog('/help', [
    function (session, args, next)
    {
        var botName = 'MAAV-1';
        var description = `I record notes and insights from the conversations Microsoft employees have with customers.  My purpose is to make it easier to disseminate knowledge about Microsoft customers and win as a team.`;

        session.send(`I'm ${botName}`);
        session.send(`${description}`);
        session.endDialog("You can enter OK to continue ...")
    }
]).triggerAction({ matches: /^help/i }); 

module.exports = { 
    connector: connector,
    bot: bot
};