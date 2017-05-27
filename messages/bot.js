"use strict";
var builder = require("botbuilder");
var botbuilder_azure = require("botbuilder-azure");

var Sequelize = require('sequelize');
var Fuse = require('fuse.js');
var Connection = require('tedious').Connection;
var Request = require('tedious').Request;
var TYPES = require('tedious').TYPES;
var path = require('path');

var useEmulator = (process.env.NODE_ENV == 'development');

//Create connection to database
if (process.env.DB_SERVER)
{
    var db_connection = new Sequelize(process.env.DB_NAME, process.env.DB_ADMIN, process.env.DB_PASSWORD,
    {
        dialect: 'mssql',
        host: process.env.DB_SERVER,
        port: 1433,
        logging: false,

        dialectOptions: {
            requestTimeout: 30000, //timeout = 30 seconds
            encrypt: true
        }
    });

    // Check database connection
    db_connection.authenticate()
        .then(function(){console.log("Connection to '" + process.env.DB_NAME + "' database has been established successfully.");})
        .catch(function (err){console.log("Unable to connect to the database:", err);})
        .done();
   
    //Define 'feedback' model
    var Feedback = db_connection.define('feedback',{
        Name: Sequelize.STRING,
        Authors: Sequelize.STRING,
        Company: Sequelize.STRING,
        Contact: Sequelize.STRING,
        Product: Sequelize.STRING,
        Notes: Sequelize.STRING,
        Summary: Sequelize.STRING,
        Tags: Sequelize.STRING
    });

    db_connection.sync().then(function()
    {
        console.log("Created database shema from 'feedback' model");
    });
}

// Create connection to chat bot
var connector = useEmulator ? new builder.ChatConnector() : new botbuilder_azure.BotServiceConnector({
    appId: process.env['MicrosoftAppId'],
    appPassword: process.env['MicrosoftAppPassword'],
    stateEndpoint: process.env['BotStateEndpoint'],
    openIdMetadata: process.env['BotOpenIdMetadata']
});

// Create chat bot instance
var bot = new builder.UniversalBot(connector);
bot.localePath(path.join(__dirname, './locale'));

// Wake bot up when a user joins the channel
bot.on('conversationUpdate', function(message)
{
    if (message.membersAdded)
    {
        message.membersAdded.forEach(function(identity)
        {
            if (identity.id === message.address.bot.id)
                bot.beginDialog(message.address, '/');
        });
    }
});

// Main dialog loop - this is the default dialog that launches other dialogs
bot.dialog('/', [
    function (session, args, next)
    {
        // Register global dialog handlers
        bot.beginDialogAction('Help', '/help');
        bot.beginDialogAction('Interactive Entry', '/interactiveDataEntry');
        bot.beginDialogAction('Batch Entry', '/batchDataEntry');
        bot.beginDialogAction('Retrieve', '/fetchConversation');
        bot.beginDialogAction('Dashboard', '/viewDashboard');

        var botChannel = session.message.address.channelId;

        // Send bot intro if this is the user's first interaction with bot
        if ((!session.userData.name) && (botChannel === "emulator"))
            session.beginDialog('/sayHello');
        else if ((!session.userData.name) && (session.message.timestamp))
        {
            var userDetails = session.message.address.user;

            if ((botChannel === "emulator") || (botChannel === "webchat"))
                session.beginDialog('/sayHello');
            else if (botChannel === "email")
            {
                session.userData.name = userDetails.id
                session.userData.firstName = userDetails.id.split('@')[0];
                session.beginDialog('/firstRun');
            }
            else if (userDetails.name)
            {
                session.userData.name = userDetails.name
                session.userData.firstName = userDetails.name.split(' ')[0];
                session.beginDialog('/firstRun');
            }
            else
                session.beginDialog('/sayHello');
        }
        // Send available bot actions if the user has previous experience with the bot       
        else if (session.userData.name)
        {
            session.beginDialog('/selectAction');   
        }
    }
]);

bot.dialog('/sayHello',
[
    function (session, args, next)
    {
        if (!session.userData.name)
        {
            session.send("Hello there! It appears that I don't know your name.");
            builder.Prompts.text(session, "What do I call you?");
        }
        else
        {
            var prompt = `Hello there! My sensors tell me your name is ${session.userData.name}`;
            prompt+= `What would you like me to call you?`;
            builder.Prompts.text(session, prompt);
        }
    },
    function (session, results)
    {
        session.userData.name = results.response;
        session.userData.firstName = results.response.split(' ')[0];
        session.replaceDialog('/firstRun');
    }
]);

//First run dialog - this dialog is run the first time a user interacts with a bot
bot.dialog('/firstRun',
[
    function (session, args, next)
    {
        var botName = 'MAAV-2';
        var description = 'I record notes and insights from the conversations Microsoft employees have with customers.  My purpose is to make it easier to disseminate knowledge about Microsoft customers and win as a team.';
        var userName = session.userData.firstName;
 
        //Send bot intro and template for email channel
        if (session.message.address.channelId === "email")
        {
            // Send markdown version of conversation card for email channel
            session.conversationData.displayMarkdown = true;
            
            // Parse email chain if bot was forwarded email
            if (isEmail(session.message.text))
            {
                var message = `Greetings ${userName},\n\n`;

                message+= `I'm ${botName}. ${description} \n\n\n\n`;
                message+= "I see you've sent a previous conversation in the email body. Give me a few minutes to process this information ...";
                session.send(message);
                session.replaceDialog('/batchParser', session.message.text);
            }
            else
            {    
                var message = `Greetings ${userName},\n\n`;

                message+= `I'm ${botName}. ${description} \n\n\n\n`;
                message+= "Below you will find the template I need to record your conversation.\n\n"
                message+= "Reply back with the completed template to continue:\n\n---\n\n";
                message+= "**Authors:** {Microsoft alias}, {Microsoft alias} ... \n\n";
                message+= "**Company:** {company name} \n\n";
                message+= "**Contact:** {customer contact name}, {customer contact name} ... \n\n";
                message+= "**Product:** {SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other} \n\n";
                message+= "**Tags (optional):** {tag}, {tag} ... \n\n";
                message+= "**Summary (optional):** {enter short summary of note here}";
                message+= "**Notes:** {enter note text here} \n\n\n\n";

                message+= "Here's an example completed template:\n\n";
                message+= "Authors: madhuka, ayolubek, anshrest, vinsonyu \n\n";
                message+= "Company: Wonka Chocolate Factory \n\n";
                message+= "Contact: Willy Wonka \n\n";
                message+= "Product: SQL VM, SQL DB \n\n";
                message+= "Tags: chocolate, column-store \n\n";
                message+= "Summary: Mr. Wonka described the challenges he's facing migrating from SQLVM to SQLDB \n\n";
                message+= "Notes: Mr. Wonka is the eccentric owner of the world-famous Wonka chocolate factory. He employees mystery workers called Oompa-Loompas to operate his chocolate factory. The Oompa-Loompas use SQL VM to crank out batches of chocolate for the millions of customers of Wonka chocolate. Their massive growth has led them to worry about how much time they're spending on maintaining infrastructure"
                message+= "instead of making delicious chocolates. They are interested in Azure SQL database but are concerned about the security implications of a public endpoint. The secret chocolate recipe algorithm cannot fall into the wrong hands or it will spell disaster for the company."

                builder.Prompts.text(session, message);
            }
        }
        // Send bot intro and option buttons for chat messenger channels
        else
        {
            session.send(`Greetings ${userName}! I'm ${botName}`);
            session.send(`In a nutshell, here's what I do:\n\n${description}`);
            var optionButtons = new builder.ThumbnailCard(session)
            .title("Bot actions")
            .buttons([
                builder.CardAction.dialogAction(session, "Help","", "Help"),
                builder.CardAction.dialogAction(session, "Interactive Entry", "", "Add Conversation (Interactive mode)"),
                builder.CardAction.dialogAction(session, "Batch Entry", "", "Add Conversation (Batch mode)"),
                builder.CardAction.dialogAction(session, "Retrieve", "", "Retrieve conversation(s)"),
                builder.CardAction.dialogAction(session, "Dashboard", "", "View conversation dashboard")
            ]);
            
            session.send("Select one of the actions below to continue:");
            session.send(new builder.Message(session).addAttachment(optionButtons));
        }
    },
    function (session, results)
    {
        //Store email response template
        var inputTemplate = results.response;
        session.beginDialog('/batchParser', inputTemplate);
    }
   
]);

// Selection dialog - this dialog presents the user with bot actions. This is typically the screen users will start with
bot.dialog('/selectAction', 
[
    function(session, args, next)
    {
        var userName = session.userData.firstName;
        var message;

        if (session.message.address.channelId === "email")
        {
            // Send markdown version of conversation card for email channel
            session.conversationData.displayMarkdown = true;

            // Parse email chain if bot was forwarded email
            if (isEmail(session.message.text))
            {
                var message = `Greetings ${userName},\n\n`;

                message+= "I see you've sent a previous conversation in the email body. Give me a few minutes to process this information ...";
                session.send(message);
                session.replaceDialog('/batchParser', session.message.text);
            }
            else
            {
                var dashboardURL = process.env.DashboardUrl;
                message = `Greetings ${userName}, \n\n`;

                message+= `I'm guessing you have a new customer conversation for me. If you're looking to see the conversations already stored in the archive instead, you can [take a look at the customer conversation dashboard](${dashboardURL}). \n\n`
                message+= "Below you will find the template I need to record your conversation.\n\n"
                message+= "Reply back with the completed template to continue:\n\n---\n\n";
                message+= "**Authors:** {Microsoft alias}, {Microsoft alias} ... \n\n";
                message+= "**Company:** {company name} \n\n";
                message+= "**Contact:** {customer contact name}, {customer contact name} ... \n\n";
                message+= "**Product:** {SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other} \n\n";
                message+= "**Tags (optional):** {tag}, {tag} ... \n\n";
                message+= "**Summary (optional):** {enter short summary of note here}\n\n";
                message+= "**Notes:** {enter note text here} \n\n\n\n";

                builder.Prompts.text(session, message);
            }
        }
        else
        {
            message = `Greetings ${userName}!\n\n Select one of the actions below to continue:`;
            var optionButtons = new builder.ThumbnailCard(session)
            .title("Bot actions")
            .buttons([
                builder.CardAction.dialogAction(session, "Help","", "Help"),
                builder.CardAction.dialogAction(session, "Interactive Entry", "", "Add Conversation (Interactive mode)"),
                builder.CardAction.dialogAction(session, "Batch Entry", "", "Add Conversation (Batch mode)"),
                builder.CardAction.dialogAction(session, "Retrieve", "", "Retrieve conversation(s)"),
                builder.CardAction.dialogAction(session, "Dashboard", "", "View conversation dashboard")
            ]);
            
            session.send(message);
            session.send(new builder.Message(session).addAttachment(optionButtons));
        }
    },
    function (session, results)
    {
        //Store email response template
        var inputTemplate = results.response;
        session.beginDialog('/batchParser', inputTemplate);
    }

]);

// Enter conversation details interactively with dialog
bot.dialog('/interactiveDataEntry',
[
    function (session, args, next)
    {
        // This forwards the button click for product selection to the appropriate handler
        if (session.message.value)
            next();
        else
            builder.Prompts.text(session, "Please enter the Microsoft aliases (separated by commas) of all who were on the call.");
    },
    function (session, results, next)
    {
        if (session.message.value)
            next();
        else
        {
            session.conversationData.authors = results.response;
            builder.Prompts.text(session, "What company did you speak with?");
        }
    },
    function (session, results, next)
    {
        if (session.message.value)
            next();
        else
        {
            session.conversationData.company = results.response;
            
            if (process.env.DB_SERVER)
            {
                //search for similar companies
                // Create connection to database
                var config = {
                userName: process.env.DB_ADMIN, // update me
                password: process.env.DB_PASSWORD, // update me
                server: process.env.DB_SERVER,
                    options: {
                        database: process.env.DB_NAME,
                        rowCollectionOnRequestCompletion: true,
                        encrypt: true
                    }
                }

                var connection = new Connection(config);

                // Attempt to connect and execute queries if connection goes through
                connection.on('connect', function(err) {
                if (err) {
                    console.log(err);
                } else {
                    console.log('Connected');
                    console.log('Reading rows from the Table...');
                    // Read all rows from table
                    var request = new Request(
                    'SELECT DISTINCT company FROM feedbacks;',
                    function(err, rowCount, rows) {
                        if (err) {
                        console.log(err);
                        } else {
                        var companies = rows.map(row => { return { name: row[0].value } });
                        console.log(rowCount + ' row(s) returned');
                        // console.log(rows);
                        var options = {
                            shouldSort: true,
                            includeMatches: true,
                            minMatchCharLength: 1,
                            keys: [
                                "name"
                            ]
                        };
                        var fuse = new Fuse(companies, options); // "list" is the item array
                        var result = fuse.search(session.conversationData.company);
                        console.log(companies);
                        console.log(result);
                        var matches=result.map(function(entry) {return entry.item.name;})
                        builder.Prompts.text(session, "I found these companies that are similar to " + session.conversationData.company + "did you mean any of these companies? " + matches + ". please type in a company name.");
                        connection.close();
                        }
                    });
                    // Execute SQL statement
                    connection.execSql(request);
                }
                });
            }
            else
                next();
        }
    },
    function (session, results, next)
    {
        if (session.message.value)
            next();
        else
        {
            if (process.env.DB_SERVER)
                session.conversationData.company = results.response;
            builder.Prompts.text(session, "Who did you speak with at " + session.conversationData.company + "?");
        }
    },
    function (session, results, next)
    {
        if (session.message.value)
            next();
        else
        {
            session.conversationData.contact = results.response;
            var productSelection = new builder.Message(session)
                .addAttachment({
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content:
                    {
                        type: "AdaptiveCard",
                        body:
                        [
                            {
                                "type": "TextBlock",
                                "text": "Which product(s) does your customer use?"
                            },
                            {
                                "type": "Input.ChoiceSet",
                                "id": "productSelection",
                                "isMultiSelect": true,
                                "style": "compact",
                                "choices": 
                                [
                                    {
                                      "title": "SQL VM",
                                      "value": "SQL VM"
                                    },
                                    {
                                        "title": "SQL DB",
                                        "value": "SQL DB"
                                    },
                                    {
                                        "title": "SQL DW",
                                        "value": "SQL DW"
                                    },
                                    {
                                        "title": "Elastic pool",
                                        "value": "Elastic pool"
                                    },
                                    {
                                        "title": "On-prem SQL Server",
                                        "value": "On-prem SQL Server"
                                    },
                                    {
                                        "title": "Other",
                                        "value": "Other"
                                    }
                                ]
                            }           
                        ],
                        actions:
                        [
                            {
                                "type": "Action.Submit",
                                "title": "Submit",
                                "data": { "message" : "productSelection"}  
                            }
                        ]
                    }
                });

                session.send(productSelection);
                next();
            }
    },
    function (session, results)
    {
        if (session.message.value)
        {
            session.conversationData.product = session.message.value.productSelection;
            builder.Prompts.text(session, "Please paste in below any notes you took down during the call");
        }
    },
    function (session, results)
    {
        session.conversationData.notes = results.response;
        session.send("Any tags you would like to add to make this conversation easier to find in the future?")
        builder.Prompts.text(session, "Popular options include tagging features your customer uses, tagging private preview status, tagging blockers and happy moments that were revealed during the conversation");
    },
    function (session, results)
    {
        session.conversationData.tags = results.response;
        session.send("That's all I need. Thanks for the info.");
        
        session.beginDialog('/displayConversationCard', session.conversationData);
    } 
]);

// Enter conversation details at once using provided template
bot.dialog('/batchDataEntry',
[
    function(session, args, next)
    {
        var userName = session.userData.firstName;
        var message = "Below you will find the template I need to record your conversation.\n\n"
        message+= "Reply back with the completed template to continue:\n\n---\n\n";
        message+= "**Authors:** {Microsoft alias}, {Microsoft alias} ... \n\n";
        message+= "**Company:** {company name} \n\n";
        message+= "**Contact:** {customer contact name}, {customer contact name} ... \n\n";
        message+= "**Product:** {SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other} \n\n";
        message+= "**Tags (optional):** {tag}, {tag} ... \n\n";
        message+= "**Summary (optional):** {enter short summary of note here}\n\n";
        message+= "**Notes:** {enter note text here} \n\n";

        builder.Prompts.text(session, message);
    },
    function(session, results)
    {
        session.replaceDialog('/batchParser', results.response);
    }
]);

// Parse conversation details that were input via batch template or email
bot.dialog('/batchParser',
[
    function(session, args, next)
    {
        // Parse input email chain
        if (isEmail(session.message.text))
        {
            // Email parser function goes here
            session.conversationData.notes = session.message.text;
        }
        // Parse input conversation template
        else
        {
            session.sendTyping();
            var emailSignatureRegex = /(^[\s]*--*[\s]*[a-z \.]*\w+$|^[\s]*best[\s,!\w]*\w+$|^[\s]*regards[\s,!\w]*\w+$|^[\s]*thanks[\s,!\w]*\w+$|^[\s]*cheers[\s,!\w]*\w+$|^[\s]*sent from [\w' ]+$)/im
            var conversationTemplateRegex = /(\w+:\s*)/i

            // Parse email signatures out of input text
            var templateTokens = args.split(emailSignatureRegex)[0];
            templateTokens = templateTokens.split(conversationTemplateRegex);

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
                else if (templateTokens[token].search(/summary:/i) != -1)
                    session.conversationData.summary = templateTokens[token+1];
            }
        }

        if (session.conversationData.displayMarkdown)
            session.beginDialog('/displayMarkdownConversationCard', session.conversationData);
        else
            session.beginDialog('/displayConversationCard', session.conversationData);
    }
]);

// Retrieve last 10 conversations for an input company
bot.dialog('/fetchConversation',
[
    function(session)
    {
        session.send('Function not yet implemented').endDialog();
    }
]);

// Launch PowerBI dashboard Url
bot.dialog('/viewDashboard',
[
   function(session)
   {
        var dashboardURL = process.env.DashboardUrl;
        if (dashboardURL)
            session.send(`The conversation dashboard is available at: ${dashboardURL}`).endDialog();
        else
            session.send(`Oh no! Looks like I wasn't configured with the location of the conversation dashboard.`).endDialog();
    }
]);

// Display conversation details using adaptive card
bot.dialog('/displayConversationCard',
[
    function(session, args, next)
    {
        var conversationObject;
        
        if (args)
            conversationObject = args;
        else
            conversationObject = new Object();

        if (!conversationObject.contact)
            conversationObject["contact"]=" ";
        if (!conversationObject.product)
            conversationObject["product"]=" ";
        if (!conversationObject.company)
            conversationObject["company"]=" ";
        if (!conversationObject.authors)
            conversationObject["authors"]=" ";
        if (!conversationObject.tags)
            conversationObject["tags"]=" ";
        if (!conversationObject.notes)
            conversationObject["notes"]=" ";

        var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + conversationObject.contact + " today where you discussed about how " + conversationObject.product + " is used at " + conversationObject.company + "</s><voice gender = \"female\"></voice>"
        var header = "Conversation with " + conversationObject.company
        
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
                        "weight": "bolder",
                        "separation": "strong",
                        "wrap": "true"
                    },
                    {
                        "type": "TextBlock",
                        "text": `**Company:**\n\n ${conversationObject.company}`,
                        "wrap": "true"
                    },
                    {
                        "type": "TextBlock",
                        "text": `**Author(s):**\n\n ${conversationObject.authors}`,
                        "separation": "strong",
                        "wrap": "true"
                    },
                    {
                        "type": "TextBlock",
                        "text": `**Customer contact(s):**\n\n ${conversationObject.contact}`,
                        "separation": "strong",
                        "wrap": "true"
                    },
                    {
                        "type": "TextBlock",
                        "text": `**Product(s) discussed:**\n\n ${conversationObject.product}`,
                        "separation": "strong",
                        "wrap": "true"
                    },
                    {
                        "type": "TextBlock",
                        "text": `**Tags:**\n\n ${conversationObject.tags}`,
                        "separation": "strong",
                        "wrap": "true"
                    },
                    {
                        "type": "TextBlock",
                        "text": `**Notes:**\n\n ${conversationObject.notes}`,
                        "separation": "strong",
                        "wrap": "true"
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

        if (!session.message.value)
        {
            session.send("Please confirm the information below is accurate:");
            session.send(outputCard);
        }
        
        next();
    },
    function (session, results)
    {
        if ((session.message) && (session.message.value))
        {
            var response = session.message.value;
            
            if (response.message === "edit")
            {
                var inputText = session.message.value.comment;
                session.message.value = null;
                session.replaceDialog('/editConversation', inputText);
            }
            else if (response.message === "confirm")
            {
                session.message.value = null;
                session.replaceDialog('/confirm');                
            }
            else
            {
                session.message.value = null;
                session.send("Did not understand that response. Please reply back with 'Edit' to edit the conversation card or 'Confirm' to confirm the card details.");
            }    
        }
    }
]);

// Display conversation details using markdown text
bot.dialog('/displayMarkdownConversationCard',
[
    function (session, args, next)
    {
        var conversationObject;
        var outputMessage;
        
        if (args)
            conversationObject = args;
        else
            conversationObject = new Object();

        if (!conversationObject.contact)
            conversationObject["contact"]=" ";
        if (!conversationObject.product)
            conversationObject["product"]=" ";
        if (!conversationObject.company)
            conversationObject["company"]=" ";
        if (!conversationObject.authors)
            conversationObject["authors"]=" ";
        if (!conversationObject.tags)
            conversationObject["tags"]=" ";
        if (!conversationObject.notes)
            conversationObject["notes"]=" ";
        if (!conversationObject.summary)
            conversationObject["summary"]=" ";

        // Confirmation and edit instructions
        outputMessage = "**Please confirm the information below is accurate**\n";
        outputMessage += "* Reply with **Confirm** to accept the conversation details below.\n";
        outputMessage += "* Reply with **Discard** to discard the conversation.\n";
        outputMessage += "* **Edit the details below** and reply if you would like to change any conversation detail below.\n\n\n\n";

        // Conversation details
        outputMessage += `**Conversation with ${conversationObject.company}**\n\n`;
        outputMessage += `>COMPANY:\n\n>${conversationObject.company}\n---\n`;
        outputMessage += `>AUTHOR(S):\n\n>${conversationObject.authors}\n---\n`;
        outputMessage += `>CUSTOMER CONTACT(S):\n\n>${conversationObject.contact}\n---\n`;
        outputMessage += `>PRODUCT(S) DISCUSSED:\n\n>${conversationObject.product}\n---\n`;
        outputMessage += `>TAGS:\n\n>${conversationObject.tags}\n---\n`;
        outputMessage += `>SUMMARY:\n\n>${conversationObject.summary}\n---\n`;
        outputMessage += `>NOTES:\n\n>${conversationObject.notes}`;
       
        builder.Prompts.text(session, outputMessage);
    },
    function (session, results)
    {
        if (/^[\s]*confirm[\s]+/im.test(session.message.text))
            session.replaceDialog('/confirm');
        else if (/^[\s]*discard[\s]+/im.test(session.message.text))
        {
            session.send("Discarding conversation.");
            session.endConversation();
        }
        else
            session.replaceDialog('/editConversation', session.message.text);
    }
]);

// Edit existing conversation details
bot.dialog('/editConversation', [
    function (session, args, next)
    {
        session.send("Updating conversation details...");

        if ((!args) && (!session.conversationData.displayMarkdown))
        {
            session.send("I did not receive any parameters to change. Resending conversation card...");
            session.replaceDialog('/displayConversationCard', session.conversationData);
        }
        else if ((!args) && (session.conversationData.displayMarkdown))
        {
            session.send("I did not receive any parameters to change. Resending conversation card...");
            session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);
        }
        else
            session.replaceDialog('/batchParser', args);
    }
]);

// End dialog
bot.dialog('/confirm', [
    function (session, args, next)
    {
     
        session.sendTyping();

        //Write feedback to database (if connected)
        if (process.env.DB_SERVER)
        {
            Feedback.create({
                Name: session.userData.name,
                Authors: session.conversationData.authors,
                Company: session.conversationData.company,
                Contact: session.conversationData.contact,
                Product: session.conversationData.product,
                Notes: session.conversationData.notes,
                Summary: session.conversationData.summary,
                Tags: session.conversationData.tags
            }).then(feedback => {console.log(feedback.get({plain: true}))});
            //db_connection.close();
        }

        session.send("Your conversation has been recorded.")
        session.endConversation();
    }
]);

//Help dialog
bot.dialog('/help', [
    function (session, args, next)
    {
        var botName = 'MAAV-2';
        var description = `I record notes and insights from the conversations Microsoft employees have with customers.  My purpose is to make it easier to disseminate knowledge about Microsoft customers and win as a team.`;
        var helpText = "* **/help**: To list available commands \n\n";
        helpText += "* **/home**: To restart conversation session and return to default screen \n\n";
        helpText += "* **/reset**: To reset bot to default state \n\n";

        session.send(`I'm ${botName}`);
        session.send(`${description}`);
        session.send(helpText);
        session.endDialog("You can enter OK to continue ...")
    }
]).triggerAction({
    matches: /^\/help/i, 
    onSelectAction: (session, args) => {session.beginDialog(args.action, args);}
});

//Restart dialog
bot.dialog('/home', [
    function (session, args, next)
    {
        session.replaceDialog('/selectAction');
    }
]).triggerAction({ matches: /^\/home/i }); 

//Reset dialog
bot.dialog('/reset', [
    function (session, args, next)
    {
        session.userData.name = null;
        session.userData.firstName = null;
        session.conversationData.authors = null;
        session.conversationData.company = null;
        session.conversationData.contact = null;
        session.conversationData.product = null;
        session.conversationData.notes = null;
        session.conversationData.tags = null;
        session.message.text = null;
        session.endDialog("Resetting to default state...");
        session.beginDialog('/');
    }
]).triggerAction({ matches: /^\/reset/i });

// Check if text is email
function isEmail(inputText)
{
    var hasSender = (inputText.search(/from:/i) != -1);
    var hasSendDate = (inputText.search(/sent:/i) != -1);
    var hasRecipient = (inputText.search(/to:[\s+\w]+<\w+@\w+.com>/i) != -1);
    var hasSubject = (inputText.search(/subject:/i) != -1);

    if (hasSender && hasSendDate && hasRecipient && hasSubject)
        return true;
    else
        return false;
} 

module.exports = { 
    connector: connector,
    bot: bot
};
