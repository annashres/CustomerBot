"use strict";
var builder = require('botbuilder');
var azure_builder = require('botbuilder-azure');

// var db_connection = require('tedious').Connection;
// var db_request = require('tedious').Request;
// var db_types = require('tedious').TYPES;


// Create connection to chatbot
// var Sequelize = require('sequelize');

// var userName = 'sa';
// var password = PASSWORD; 
// var hostName = 'localhost';
// var sampleDbName = DATABASE;

// var sampleDb = new Sequelize(sampleDbName, userName, password, {
//     dialect: 'mssql',
//     host: hostName,
//     port: 1433, // Default port
//     logging: false, // disable logging; default: console.log

//     dialectOptions: {
//         requestTimeout: 30000 // timeout = 30 seconds
//     }
// });

// // Define the 'Feedback' model
// var Feedback = sampleDb.define('feedback', {
//     Name: Sequelize.STRING,
//     Authors: Sequelize.STRING,
//     Company: Sequelize.STRING,
//     Contact: Sequelize.STRING,
//     Product: Sequelize.STRING,
//     Notes: Sequelize.STRING,
//     Tags: Sequelize.STRING
// });

// sampleDb.sync()
// .then(function() {
//     console.log('\nCreated database schema from model.');
// })
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});


// Create connection to database
var db_config = {
  userName: process.env.DB_ADMIN,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  options: {
      database: process.env.DB_NAME,
      encrypt: true
  }
}

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
        session.send("Nice to meet you %s.", session.userData.name); 
        session.send("How would you like to record your customer conversation?");
        session.message.text = null;
        session.beginDialog('/EntrySelection');
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
        else if (session.message.address.channelId === "email")
        {
            // Default to batch data entry for email channel
            session.beginDialog('/batchDataEntry');
        }
        else
        {
            session.send("Welcome back %s! I'm guessing you have a new customer conversation for me.", session.userData.name);
            session.message.text = null;
            session.send("How would you like to record your customer conversation?")
            session.beginDialog('/EntrySelection');
        }
    }
]);

// Entry selection dialog
bot.dialog('/EntrySelection', 
[
    function(session, args, next)
    {
        if (!session.message.text)
        {
            var dataEntrySelection = new builder.ThumbnailCard(session)
                .title("Data Entry Method")
                .buttons([
                    builder.CardAction.imBack(session, "Interactive Entry", "Interactive Entry"),
                    builder.CardAction.imBack(session, "Batch Entry", "Batch Entry")
                ]);
            session.send(new builder.Message(session).addAttachment(dataEntrySelection));
        }
        else
        {
            var selection = session.message.text;

            if (selection === 'Interactive Entry')
            {
                session.message.text = null;
                session.replaceDialog('/interactiveDataEntry')
            }
            else if (selection === 'Batch Entry')
            {
                session.message.text = null;
                session.replaceDialog('/batchDataEntry');
            }
            else
            {
                session.send("Did not understand selection '%s'", selection);
                session.message.text = null;
            }
        }
    }

]);

// Interactive entry dialog
bot.dialog('/interactiveDataEntry',
[
    function (session)
    {
        builder.Prompts.text(session, "Besides yourself, who else was on the call");
    },
    function (session, results)
    {
        session.conversationData.authors = results.response;
        builder.Prompts.text(session, "What company did y'all speak with?");
    },
    function (session, results)
    {
        session.conversationData.company = results.response;
        builder.Prompts.text(session, "Who did you speak with at " + session.conversationData.company + "?");
    },
    function (session, results)
    {
        session.conversationData.contact = results.response;
        builder.Prompts.choice(session, "Which products did you talk about?", ["SQLVM", "Elastic Pools", "Azure SQL DB", "On-Prem SQL Server", "Other"]);
    },
    function (session, results)
    {
        session.userData.product = results.response.entity;
        builder.Prompts.text(session, "Please paste in below any notes you took down during the call");
    },
    function (session, results)
    {
        session.conversationData.notes = results.response;
        builder.Prompts.text(session, "Any tags you would like to add to make this conversation easier to find in the future?");
    },
    function (session, results)
    {
        session.conversationData.tags = results.response;
        session.send("That's all I need.");
        
        session.beginDialog('/conversationCard');
    } 
]);

// Batch entry dialog
bot.dialog('/batchDataEntry',
[
    function(session)
    {
        var inputFeedback = session.conversationData.inputFeedback;
        if (!inputFeedback)
            inputFeedback = ";";
        
        if(!session.conversationData.templateDisplayed)
        {
            var feedbackTemplate = "Enter your conversation using the following template: \n\n";
            //feedbackTemplate+= "@maav1: \n\n";
            feedbackTemplate+= "**Authors:** {Microsoft alias}, {Microsoft alias} \n\n";
            feedbackTemplate+= "**Company:** {company name} \n\n";
            feedbackTemplate+= "**Contact:** {customer contact name} \n\n";
            feedbackTemplate+= "**Product:** {SQLVM, Elastic Pools, Azure SQL DB, On-Prem SQL Server, Other} \n\n";
            feedbackTemplate+= "**Tags:** {tag}, {tag} \n\n";
            feedbackTemplate+= "**Notes:** {enter note text here} \n\n";
            feedbackTemplate+= "**/END**";

            session.send(feedbackTemplate);
            session.conversationData.templateDisplayed = true;
        }
        

        //Read input until termination character
        var templateReceived = (inputFeedback.indexOf('/END') != -1);
        if (inputFeedback && (!templateReceived))
        {
            session.conversationData.inputFeedback+= session.message.text;
            session.message.text = null;
            console.log('receiving conversation details');

            inputFeedback = session.conversationData.inputFeedback.toString();
            templateReceived = (inputFeedback.indexOf('/END') != -1);
        }
        
        if (templateReceived)
        {
            console.log('conversation received');
            session.sendTyping();
            var templateTokens = session.conversationData.inputFeedback.split(/(\w+:\s*)/i)

            for (var token = 0; token<templateTokens.length; token++)
            {
                if (templateTokens[token].search(/authors?:/i) != -1)
                    session.conversationData.authors = templateTokens[token+1];
                else if (templateTokens[token].search(/company:/i) != -1)
                    session.conversationData.company = templateTokens[token+1];
                else if (templateTokens[token].search(/contacts?:/i) != -1)
                    session.conversationData.contact = templateTokens[token+1];
                else if (templateTokens[token].search(/products?:/i) != -1)
                    session.conversationData.product = templateTokens[token+1];
                else if (templateTokens[token].search(/tags?:/i) != -1)
                    session.conversationData.tags = templateTokens[token+1];
                else if (templateTokens[token].search(/notes?:/i) != -1)
                    session.conversationData.notes = templateTokens[token+1];
            }

            console.log(session.conversationData);

            session.beginDialog('/conversationCard');
        }
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
                
            var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + session.conversationData.contact + " today where you discussed about how " + session.userData.product + " is used at " + session.conversationData.company + "</s><voice gender = \"female\"></voice>"
            var header = "Conversation with " + session.conversationData.company
            var authors = session.userData.name + ", " + session.conversationData.authors

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
                                    "value": session.conversationData.company
                                },
                                {
                                    "title": "Author(s):",
                                    "value": authors
                                },
                                {
                                    "title": "Customer contact(s):",
                                    "value": session.conversationData.contact
                                },
                                {
                                    "title": "Product(s) discussed:",
                                    "value": session.conversationData.product
                                },
                                {
                                    "title": "Tags:",
                                    "value": session.conversationData.tags
                                },
                                {
                                    "title": "Notes:",
                                    "value": session.conversationData.notes
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
                // Feedback.create({Name: session.userData.name, Authors: session.conversationData.authors, Company: session.conversationData.company, Contact: session.conversationData.contact, Product: session.userData.product, Notes: session.conversationData.notes, Tags: session.conversationData.tags}).then(feedback => {
                // console.log(feedback.get({ plain: true}))
                // });
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
        
        var feedbackObject = new Object();
        feedbackObject.name = session.userData.name;
        feedbackObject.authors = session.conversationData.authors;
        feedbackObject.company = session.conversationData.company;
        feedbackObject.contact = session.conversationData.contact;
        feedbackObject.product = session.userData.product;
        feedbackObject.notes = session.conversationData.notes;
        feedbackObject.tags = session.conversationData.tags;

        //Write feedback to database if available
        if (db_config.server)
        {
            // Create connection to database if available
            var db_session = new db_connection(db_config);

            // Connect 
            db_session.on('connect', function(err) {
                if (err) {
                    console.log(err)
                }
                else{
                    console.log('connected')
                }
            });

            // Construct SQL query
            var db_query = new db_request('createFeedback', function(err)
            {
                if (err){ console.log(err); }

            });
            db_query.addParameter('feedbackObject', TYPES.NVarChar, feedbackObject);

            //Exedute query
            db_session.callProcedure(db_query);
            db_session.close();
                
        }

        session.send("Your conversation has been recorded. Bye for now")
        session.endDialog();
    }
]);

//Help dialog
bot.dialog('/help', [
    function (session, args, next)
    {
        var botName = 'MAAV-1';
        var description = `I record notes and insights from the conversations Microsoft employees have with customers.  My purpose is to make it easier to disseminate knowledge about Microsoft customers and win as a team.`;
        var helpText = "* To list available commands: **/help** \n\n";
        helpText += "* To restart conversation session: **/restart** \n\n";
        helpText += "* To reset bot to default state: **/reset** \n\n";

        session.send(`I'm ${botName}`);
        session.send(`${description}`);
        session.send(helpText);
        session.endDialog("You can enter OK to continue ...")
    }
]).triggerAction({ matches: /^\/help/i }); 

//Restart dialog
bot.dialog('/restart', [
    function (session, args, next)
    {
        session.endDialog("Restarting conversation session ...");
        session.beginDialog('/');
    }
]).triggerAction({ matches: /^\/restart/i }); 

//Reset dialog
bot.dialog('/reset', [
    function (session, args, next)
    {
        session.userData.name = null;
        session.conversationData.authors = null;
        session.conversationData.company = null;
        session.conversationData.contact = null;
        session.userData.product = null;
        session.conversationData.notes = null;
        session.conversationData.tags = null;
        session.userData.inputFeedback = null;
        session.message.text = null;
        session.endDialog("Resetting to default state...");
        session.beginDialog('/');
    }
]).triggerAction({ matches: /^\/reset/i }); 

module.exports = { 
    connector: connector,
    bot: bot
};