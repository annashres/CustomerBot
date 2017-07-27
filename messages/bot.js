"use strict";
var builder = require("botbuilder");
var botbuilder_azure = require("botbuilder-azure");

var Sequelize = require('sequelize');
var Fuse = require('fuse.js');
var Feedback;
var app_authcode;
// var Connection = require('tedious').Connection;
// var Request = require('tedious').Request;
// var TYPES = require('tedious').TYPES;
var path = require('path');
var dbconnection = require("seriate");
var botdisplay = require('./botdisplay.js');
var useEmulator = (process.env.NODE_ENV == 'development');

//Create connection to database
if (process.env.DB_SERVER)
{
    // var db_connection = new Sequelize(process.env.DB_NAME, process.env.DB_ADMIN, process.env.DB_PASSWORD,
    // {
    //     dialect: 'mssql',
    //     host: process.env.DB_SERVER,
    //     port: 1433,
    //     logging: false,

    //     dialectOptions: {
    //         requestTimeout: 30000, //timeout = 30 seconds
    //         encrypt: true
    //     }
    // });

    // Check database connection
    // db_connection.authenticate()
    //     .then(function(){console.log("Connection to '" + process.env.DB_NAME + "' database has been established successfully.");})
    //     .catch(function (err){console.log("Unable to connect to the database:", err);})
    //     .done();
   
    //Define 'feedback' model
    // Feedback = db_connection.define('feedback',{
    //     Name: Sequelize.STRING,
    //     Authors: Sequelize.STRING,
    //     Company: Sequelize.STRING,
    //     Contact: Sequelize.STRING,
    //     Product: Sequelize.STRING,
    //     Notes: Sequelize.TEXT,
    //     Summary: Sequelize.STRING(400),
    //     Tags: Sequelize.STRING,
    //     Blockers: Sequelize.STRING,
    //     ProjectStage: Sequelize.STRING
    // });

    // db_connection.sync().then(function()
    // {
    //     console.log("Created database schema from 'feedback' model");
    // });

    // Create connection to database
    var config = 
    {
        "user": process.env.DB_ADMIN,
        "password": process.env.DB_PASSWORD,
        "server": process.env.DB_SERVER,
        "database": process.env.DB_NAME,
        "requestTimeout": 300000,
        "connectionTimeout": 300000,
        
        "options": {"encrypt": true}
    }

    var feedbackConfig = 
    {
        "name": "feedbackDb",
        "user": process.env.FEEDBACKDB_ADMIN,
        "password": process.env.FEEDBACKDB_PASSWORD,
        "server": process.env.FEEDBACKDB_SERVER,
        "database": process.env.FEEDBACKDB_NAME,
        "requestTimeout": 300000,
        "connectionTimeout": 300000,

        "options": {"encrypt": true}
    }

    dbconnection.setDefaultConfig(config);
    dbconnection.addConnection(feedbackConfig);
    console.log("Connection pool created for '" + process.env.DB_NAME + "' database.");

    //Initialize email table that is used by the bot
    var createEmailTableQuery = `IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='customer_emaildomain' and xtype='U')
    CREATE TABLE customer_emaildomain (
        ms_customer_guid UNIQUEIDENTIFIER NOT NULL,
        email_domain NVARCHAR(50) NOT NULL,
        CONSTRAINT uemaildomains UNIQUE NONCLUSTERED ( ms_customer_guid, email_domain)
    )`;

    dbconnection.execute('feedbackDb',{
        query: createEmailTableQuery
    }).then (function (results)
        {
            console.log('Email table initialized');
        }, function (err)
        {
            console.error(`Could not create customer email table`, err);
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
        var botChannel = session.message.address.channelId;
        session.conversationData.pinExists = false;
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
                session.userData.alias = session.userData.firstName;
                session.beginDialog('/firstRun');
            }
            else if (userDetails.name)
            {
                session.userData.name = userDetails.name
                session.userData.firstName = userDetails.name.split(' ')[0];
                session.beginDialog('/sayHello');
            }
            else
                session.beginDialog('/sayHello');
        }
        // Send available bot actions if the user has previous experience with the bot       
        else if (session.userData.name && botChannel != "email")
        {
            session.userData.previousExperience = true;
            session.beginDialog('/auth');   
        }
        else if (session.userData.name && botChannel === "email")
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
        else if (!session.userData.alias)
        {
            var prompt = `Hello there! My sensors tell me your name is ${session.userData.name}\n\n`;
            prompt+= `What is your Microsoft alias?`;
            builder.Prompts.text(session, prompt);
        }
    },
    function (session, results)
    {
        if (!session.userData.name)
        {
            session.userData.name = results.response;
            session.userData.firstName = results.response.split(' ')[0];
            builder.Prompts.text(session, `What is your Microsoft alias?`);
        }
        else
        {
            session.userData.alias = results.response;
            session.replaceDialog('/auth');
        }
    },
    function (session, results)
    {
        session.userData.alias = results.response;
        session.replaceDialog('/auth');
    }
]);

bot.dialog('/auth',
[
    function (session, args) {
        session.send('[Click to get your pin](https://customerauthbot.azurewebsites.net/)');
        var prompt = "Please type your six-digit pin";
        builder.Prompts.text(session, prompt);

    },
    function (session, results) {
        session.dialogData.pin = results.response;
        var email = session.userData.alias
        email += '@microsoft.com' 

        var sqlAuthQuery = `SELECT TOP 1 Code FROM [dbo].[app_authcodes]  WHERE Alias='${email}'`;
        dbconnection.execute({
            query: sqlAuthQuery
        }).then (function (results)
        {
            if (results[0].Code == session.dialogData.pin) {
                session.conversationData.pinExists = true;
               
               if (session.userData.previousExperience == true)
               {
                   session.replaceDialog('/selectAction');
               }
               else {
                session.replaceDialog('/firstRun');
               }

            }
            else{
                session.send("Your alias and pin do not match");   
                session.replaceDialog('/auth');
            }
        }, function (err)
        {
            console.error(`Could not retrieve stored conversations for ${session.dialogData.inputCompany}:`, err);
        });   
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
                var message = `Greetings ${userName},  \n\n`;

                message+= `I'm ${botName}. ${description}  \n\n`;
                message+= "I see you've sent a previous conversation in the email body. Give me a few minutes to process this information ...";
                session.send(message);
                session.replaceDialog('/batchParser', session.message.text);
            }
            else
            {    
                var defaultTemplate = getConversationTemplate();
                var message = `Greetings ${userName},\n\n  `;

                message+= `I'm ${botName}. ${description} \n\n`;
                message+= "`---`\n\n";
                message+= "Below you will find the template I need to record your conversation.\n\n"
                message+= "Reply back with the completed template to continue:\n\n";
                message+= "`---`\n\n";
                message+= defaultTemplate;
                message+= "`---`\n\n";

                message+= "*Looking for an example completed template? [You can find one here](https://raw.githubusercontent.com/annashres/CustomerBot/master/ExampleTemplate.txt)*\n\n";
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
                builder.CardAction.imBack(session, "Help", "Help"),
                builder.CardAction.imBack(session, "Interactive Entry", "Add Conversation (Interactive mode)"),
                builder.CardAction.imBack(session, "Batch Entry", "Add Conversation (Batch mode)"),
                builder.CardAction.imBack(session, "Retrieve Entry", "Retrieve conversation(s)"),
                builder.CardAction.imBack(session, "Dashboard", "View conversation dashboard")
            ]);
            
            session.send("Select one of the actions below to continue:");
            session.send(new builder.Message(session).addAttachment(optionButtons));
        }
    },
    function (session, results)
    {
        //Store email response template
        if (session.message.address.channelId === "email")
        {
            var inputTemplate = results.response;
            session.beginDialog('/batchParser', inputTemplate);
        }
        else
            session.replaceDialog('/selectAction');
    }
   
]);

// Selection dialog - this dialog presents the user with bot actions. This is typically the screen users will start with
bot.dialog('/selectAction', 
[
    function(session, args, next)
    {
        if (session.userData.alias === 'vinsonyu'){
            var message = `Hello Vin💩`;
            session.send(message);
        }

        if (session.message.address.channelId != "email" && session.conversationData.pinExists == false) {
            session.replaceDialog('/auth');
        }
        var userName = session.userData.firstName;
        var message;

        console.log(session.message);
        if (session.message.address.channelId === "email")
        {
            // Send markdown version of conversation card for email channel
            session.conversationData.displayMarkdown = true;

            // Parse email chain if bot was forwarded email
            if (isEmail(session.message.text))
            {
                session.replaceDialog('/batchParser', session.message.text);
            }
            // Parse conversation template if bot was sent a template as part of greeting
            else if (containsTemplateKeyword(session.message.text))
            {
                session.replaceDialog('/batchParser', session.message.text);
            }
            else
            {
                var dashboardURL = process.env.DashboardUrl;
                var defaultTemplate = getConversationTemplate();
                message = `Greetings ${userName},\n\n  `;

                message+= `I'm guessing you have a new customer conversation for me. If you're looking to see existing conversations, you can [view the conversation dashboard](${dashboardURL}). \n\n`
                message+= "`---`\n\n";
                message+= "Below you will find the template I need to record your conversation.\n\n"
                message+= "Reply back with the completed template to continue:\n\n";
                message+= "`---`\n\n";
                message+= defaultTemplate;
                message+= "`---`\n\n";
                message+= "*Looking for an example completed template? [You can find one here](https://raw.githubusercontent.com/annashres/CustomerBot/master/ExampleTemplate.txt)*\n\n";
    
                builder.Prompts.text(session, message);
            }
        }
        else
        {
            message = `Greetings ${userName}!\n\n Select one of the actions below to continue:`;
            var optionButtons = new builder.ThumbnailCard(session)
            .title("Bot actions")
            .buttons([
                builder.CardAction.imBack(session, "Help", "Help"),
                builder.CardAction.imBack(session, "Interactive Entry", "Add Conversation (Interactive mode)"),
                builder.CardAction.imBack(session, "Batch Entry", "Add Conversation (Batch mode)"),
                builder.CardAction.imBack(session, "Retrieve Entry", "Retrieve conversation(s)"),
                builder.CardAction.imBack(session, "Dashboard", "View conversation dashboard")
            ]);
            
            session.send(message);
            session.send(new builder.Message(session).addAttachment(optionButtons));
        }
    },
    function (session, results)
    {
        //Store email response template
        if (session.message.address.channelId === "email")
        {
            var inputTemplate = results.response;
            session.beginDialog('/batchParser', inputTemplate);
        }
        else
            session.replaceDialog('/selectAction');
    }

]);

// Enter conversation details interactively with dialog
bot.dialog('/interactiveDataEntry',
[
    function (session, args, next)
    {
        session.conversationData.interactiveMode = true;
        if (session.message.address.channelId != 'email')
        {
            if (session.conversationData.pinExists == false) {
                session.replaceDialog('/auth')
            }
            var forwardInput = session.dialogData.forwardInput;
            var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];

            // This forwards the button click for product selection to the appropriate handler
            if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
                next();
            else
                builder.Prompts.text(session, "Enter the Microsoft aliases (separated by commas) of all who were on the call.");
        }
    },
    function (session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];

        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else
        {
            session.conversationData.authors = results.response;
            session.beginDialog('/findCompanyMatches');
        }
    },
    function (session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];
        
        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else
        {
            builder.Prompts.text(session, "Who did you speak with at " + session.conversationData.company + "? (use commas to include multiple people)")
        }
    },
    function (session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];
        
        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else
        {
            session.conversationData.contact = results.response;
            session.send("Which SQL product(s) does the company use?");
            builder.Prompts.text(session, "Available options (use comma to enter multiple products): {SQL VM, SQL DB, SQL DW, Elastic pool, On-prem SQL Server, Other}");
        }
    },
    function (session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];
        
        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else
        {
            var inputText = results.response;
            var productSelection = "";

            if (inputText.match(/VM/i))
                productSelection += "SQL VM,";
            if (inputText.match(/DB/i))
                productSelection += "SQL DB,";
            if (inputText.match(/pool|elastic/i))
                productSelection += "Elastic pool,";
            if (inputText.match(/on-prem|server/i))
                productSelection += "On-prem SQL Server,";
            if (inputText.match(/other/i))
                productSelection += "Other,";
            productSelection = productSelection.replace(/,$/g, "");

            session.conversationData.product = productSelection;
            builder.Prompts.text(session, "Paste in below any notes you took down during the call");
        }
    },
    function (session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];
        
        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else if (typeof forwardInput != "undefined" && (currentStep == forwardInput))
        {
            if (session.message.text)
            {
                if (session.message.text == "Skip")
                    session.conversationData["projectstage"] = "{Select one of: Pre-POC, POC, Production}";
                else
                    session.conversationData["projectstage"] = session.message.text;

                session.dialogData.forwardInput = null;
                next();
            }
        }
        else
        {
            session.conversationData.notes = results.response;
            session.send("What stage of development is the company's product currently in?");
           
            var stageOptions = new builder.ThumbnailCard(session)
                .title("Project Stage options")
                .buttons([
                    builder.CardAction.imBack(session, "Skip", "Skip this question"),
                    builder.CardAction.imBack(session, "Pre-POC", "Pre-ProofOfConcept"),
                    builder.CardAction.imBack(session, "POC", "ProofOfConcept"),
                    builder.CardAction.imBack(session, "Production", "Production")
                ]);

            session.dialogData["forwardInput"] = currentStep;
            session.send(new builder.Message(session).addAttachment(stageOptions));
        }        
    },
    function (session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];
        
        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else if (typeof forwardInput != "undefined" && (currentStep == forwardInput))
        {
            if (session.message.text)
            {
                if (session.message.text == "Skip")
                    session.conversationData["blockers"] = "{Enter a comma-separated list of blockers if any}";
                else
                    session.conversationData["blockers"] = session.message.text;

                session.dialogData.forwardInput = null;
                next();
            }
        }
        else 
        {
            session.send("Are there any blockers that are preventing continuing development of the product? Enter a comma-separated list below or hit the 'Skip' button to skip this question.");

            var skipButton = new builder.ThumbnailCard(session)
                .title("Skip this question")
                .buttons([builder.CardAction.imBack(session, "Skip", "Skip")]);

            session.dialogData["forwardInput"] = currentStep;
            session.send(new builder.Message(session).addAttachment(skipButton));
        }       
    },


    function (session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];
        
        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else if (typeof forwardInput != "undefined" && (currentStep == forwardInput))
        {
            if (session.message.text)
            {
                if (session.message.text == "Skip")
                    session.conversationData["satisfaction"] = "{Select a customer satisfaction level from 1-4 with 4 being very satisfied and 1 being very dissatisfied. Enter 0 for unknown}";
                else
                    session.conversationData["satisfaction"] = session.message.text;

                session.dialogData.forwardInput = null;
                next();
            }
        }
        else
        {
            session.conversationData.notes = results.response;
            session.send("How satisfied is the customer?");
           
            var satOptions = new builder.ThumbnailCard(session)
                .title("Customer Satifaction Options")
                .buttons([
                    builder.CardAction.imBack(session, "Skip", "Skip this question"),
                    builder.CardAction.imBack(session, "4", "4 - Very Satisfied"),
                    builder.CardAction.imBack(session, "3", "3 - Somewhat Satisfied"),
                    builder.CardAction.imBack(session, "2", "2 - Somewhat Dissatisfied"),
                    builder.CardAction.imBack(session, "1", "1 - Very Dissatisfied"),
                    builder.CardAction.imBack(session, "0", "0 - Unknown")

                ]);

            session.dialogData["forwardInput"] = currentStep;
            session.send(new builder.Message(session).addAttachment(satOptions));
        }        
    },

    function(session, results, next)
    {
        var forwardInput = session.dialogData.forwardInput;
        var currentStep = session.dialogData['BotBuilder.Data.WaterfallStep'];
        
        // This forwards the button click for product selection to the appropriate handler
        if (typeof forwardInput != "undefined" && (currentStep != forwardInput) && (forwardInput != null))
            next();
        else if (typeof forwardInput != "undefined" && (currentStep == forwardInput))
        {
            if (session.message.text)
            {
                if (session.message.text == "Skip")
                    session.conversationData["tags"] = "{Enter a comma-separated list of tags if any}";
                else
                    session.conversationData["tags"] = session.message.text;

                session.dialogData.forwardInput = null;
                next();
            }
        }
        else
        {
            session.send("Are there any tags you would like to add to make this conversation easier to find in the future? Enter a comma-separated list below or hit the 'Skip' button to skip this question.");

            var skipButton = new builder.ThumbnailCard(session)
                .title("Skip this question")
                .buttons([builder.CardAction.imBack(session, "Skip", "Skip")]);

            session.dialogData["forwardInput"] = currentStep;
            session.send(new builder.Message(session).addAttachment(skipButton));
        }       
    },
    function (session, results)
    {
        session.send("That's all I need. Thanks for the info.");

        //Set properties to null if they have default values
        if (session.conversationData.projectstage == "{Select one of: Pre-POC, POC, Production}")
            session.conversationData["projectstage"] = null;
        if (session.conversationData.blockers == "{Enter a comma-separated list of blockers if any}")
            session.conversationData["blockers"] = null;
        if (session.conversationData.satisfaction == "{Select a customer satisfaction level from 1-4 with 4 being very satisfied and 1 being very dissatisfied. Enter 0 for unknown}")
            session.conversationData["satisfaction"] = null;
        if (session.conversationData.tags == "{Enter a comma-separated list of tags if any}")
            session.conversationData["tags"] = null;

        console.log(session.conversationData);
        
        if (session.message.address.channelId === "emulator")
            session.beginDialog('/displayConversationCard', session.conversationData);
        else
            session.beginDialog('/displayMarkdownConversationCard', session.conversationData);
    } 
]).triggerAction({ matches: /^Interactive Entry$/i });

// Enter conversation details at once using provided template
bot.dialog('/batchDataEntry',
[
    function(session, args, next)
    {
       if (session.conversationData.pinExists == false) {
            session.replaceDialog('/auth')
        }
        var userName = session.userData.firstName;
        var defaultTemplate = getConversationTemplate();
        var message = "Below you will find the template I need to record your conversation.\n\n"
        message+= "Reply back with the completed template to continue:\n\n---\n\n";
        message+= defaultTemplate;
        
        builder.Prompts.text(session, message);
    },
    function(session, results)
    {
        session.replaceDialog('/batchParser', results.response);
    }
]).triggerAction({ matches: /^Batch Entry$/i });

// Parse conversation details that were input via batch template or email
bot.dialog('/batchParser',
[
    function(session, args, next)
    {
        // Parse input email chain
        if ((isEmail(session.message.text)) && !(isValidTemplate(session.message.text)))
        {
            // Email parser function goes here
            
            var emailSenderList = [];
            var emailSenderRegex = /from: ([\w ]+)\[mailto:(\w+@\w+.com)\]|from: ([\w ,]+)<(.+@.+\.com)>/ig;
            var emailMatches = emailSenderRegex.exec(session.message.text);
            var msftContacts = "";
            var companyName = "";
            var companyContacts = "";
            var companyEmailDomain = "";

            // Parse out all the email senders into a list
            while (emailMatches != null)
            {
                emailMatches = emailMatches.filter(n => n);
                var author = {name: emailMatches[1].replace(',',''), email: emailMatches[2]};
                emailSenderList.push(author);
                emailMatches = emailSenderRegex.exec(session.message.text);           
            }

            for (var i=0; i<emailSenderList.length; i++)
            {
                var author = emailSenderList[i];
                
                //Extract MSFT alias
                if (author.email.includes("@microsoft.com"))
                {
                   var authorAlias = author.email.split("@microsoft.com")[0];
                   if (!msftContacts.includes(authorAlias))
                        msftContacts = msftContacts + "," + authorAlias
                }
                //Extract company name and contact info
                else
                {
                    if (!companyContacts.includes(author.name))
                    {
                        companyContacts = companyContacts + "," + author.name
                        companyName = author.email.split("@")[1].replace(/\.[\w]+/g,'');
                        companyEmailDomain = author.email.split("@")[1];
                    }
                }
            }

            // Append user's alias if it's not included in list of authors
            if (!msftContacts.includes(session.userData.alias))
                msftContacts = session.userData.alias + ", " + msftContacts;
            
            //Strip out any leading/trailing commas
            msftContacts = msftContacts.replace(/^[, ]+|[, ]+$/g, "");
            companyContacts = companyContacts.replace(/^,|,$/g, "");
            console.log(msftContacts);

            //Save extracted information into conversation variables
            session.conversationData["authors"] = msftContacts;
            //session.conversationData["company"] = companyName;
            session.conversationData["contact"] = companyContacts;
            session.conversationData["notes"] = botdisplay.renderEmailConversation(session.message.text);

            //TODO: error here too
            if (companyEmailDomain)
                session.beginDialog('/findCompanyFromEmail', companyEmailDomain);
            else

            //TODO: think the error is here
                //session.beginDialog('/findCompanyMatches', companyName);
                console.log("got here");
                session.conversationData["authors"] = msftContacts;
                session.conversationData["notes"] = botdisplay.renderEmailConversation(session.message.text);

        }
        // Parse response to email conversation template
        else if ((isEmail(session.message.text)) && (isValidTemplate(session.message.text)))
        {
            console.log('parsing email response template');
            parseConversationTemplate(session, session.message.text);
        }
        // Parse input conversation template (teams only?)
        
        else
        {
            session.sendTyping();
            parseConversationTemplate(session, args);
        }    
    },
    function (session, results, next)
    {
        if (session.conversationData.displayMarkdown)
            session.beginDialog('/displayMarkdownConversationCard', session.conversationData);
        else if (session.message.address.channelId === "emulator")
            session.beginDialog('/displayConversationCard', session.conversationData);
        else
            session.beginDialog('/displayMarkdownConversationCard', session.conversationData);   
    }
]);

// Retrieve last 10 conversations for an input company
bot.dialog('/fetchConversation',
[
    function(session, args, next)
    {
        if (session.message.address.channelId != 'email')
        {
            session.conversationData.interactiveMode = true;
            session.conversationData.company = null;
            session.conversationData.customerGuid = null;

            if (session.conversationData.pinExists == false)
                session.replaceDialog('/auth')
        
            if (session.message.text.match(/^conv#/))
            {
                var selection = session.message.text.split(':')[1].trim();
                var selectedConversation = session.conversationData.retrievedConversations[selection];
                var outputCard;
                var buttonArray;

                if (session.message.address.channelId === "emulator")
                    outputCard = botdisplay.renderCard(session, builder, selectedConversation);
                else
                    outputCard = botdisplay.renderText("", selectedConversation);

                session.send(outputCard);

                if (session.conversationData.retrievedConversations.length > 1)
                {
                    buttonArray = [
                        builder.CardAction.imBack(session, "More Conversations", "More conversations..."),
                        builder.CardAction.imBack(session, "Return Home", "Return home")
                    ];
                }
                else
                    buttonArray = [builder.CardAction.imBack(session, "Return Home", "Return home")];
                
                var optionButtons = new builder.ThumbnailCard(session)
                .title("Available actions")
                .buttons(buttonArray);

                session.send(new builder.Message(session).addAttachment(optionButtons));            
            }
            else if (session.message.text.match(/^more conversations/i))
            {
                var conversations = botdisplay.renderSummaryCard(session, builder, session.conversationData.retrievedConversations);

                var outputCards = new builder.Message(session)
                    .attachmentLayout(builder.AttachmentLayout.carousel)
                    .attachments(conversations);

                session.send(outputCards);
            }
            else if (session.message.text.match(/^return home/i))
            {
                session.conversationData.retrievedConversations = null;
                session.replaceDialog('/selectAction');
            }
            else if (process.env.DB_SERVER)
                builder.Prompts.text(session, "Which company would you like to retrieve conversations for?");
            else if (!process.env.DB_SERVER)
                session.endDialog("This function is not available. I have not been hooked up to a database of previous conversations.");   
        }      
    },
    function (session, results, next)
    {
        // Get input company to search for
        if(!session.dialogData.inputCompany)
            session.dialogData.inputCompany = results.response;

        session.beginDialog('/findCompanyMatches', session.dialogData.inputCompany);
        // Find input company in feedback database
        // var sqlCompanyQuery = "SELECT DISTINCT company FROM feedbacks;";

        // dbconnection.execute({
        //     query: sqlCompanyQuery
        // }).then (function (results)
        // {
        //     var companies = results
        //     var searchOptions = 
        //     {
        //         shouldSort: true,
        //         includeMatches: true,
        //         minMatchCharLength: 1,
        //         keys: ["company"],
        //         threshold: 0.4
        //     }
        //     var fuse = new Fuse(companies, searchOptions);

        //     // Search database for input company name and companies with matching names
        //     var result = fuse.search(session.dialogData.inputCompany);
        //     var matches = result.map(function (entry) {return entry.item.company;});
        //     console.log(matches);

        //     if (matches.length == 0)
        //     {
        //         session.send(`There is no record of '${session.dialogData.inputCompany}' in the database. Please refine your search or try another company.`);
        //         session.endDialog("Enter OK to return to home screen and try again.");
        //     }    
        //     else if (matches.length > 1)
        //     {
        //         session.send(`I found a few companies that match '${session.dialogData.inputCompany}`);
        //         var prompt = "Please select one company from the choices below:\n\n";
        //         builder.Prompts.choice(session, prompt, matches);
        //     }
        //     else if (matches.length == 1)
        //     {
        //         var prompt = `Found matching company '${matches}'. Retrieving last 10 conversations...`;
        //         session.send(prompt);
        //         session.dialogData.inputCompany = matches;
        //         next();
        //     }
        // }, function (err)
        // {
        //     console.error("Could not search database for provided company name:", err);
        // });
    },
    function (session, results)
    {
        // if (results.response)
        //     session.dialogData.inputCompany = results.response.entity;
        var conversationListQuery = `SELECT TOP 10 [conversation_id] AS id, [creator_alias] AS [Name],[authors],[company], [customer_contacts] AS [Contact], [service_discussed_text] AS [Product],[notes],[summary], [search_tags] AS [Tags], [blocker_tags] AS [Blockers], [project_stage_tags] AS [ProjectStage], [satisfaction] AS [Satisfaction], [update_time] AS [updatedAt]
                                     FROM [dbo].[conversationDetails] 
                                     WHERE [ms_customer_guid] = '${session.conversationData.customerGuid}'
                                     ORDER BY [updatedAt] DESC`;

        dbconnection.execute('feedbackDb', {
            query: conversationListQuery
        }).then (function (results)
        {
            if (results.length)
            {
                var initialOutput = `Found matching company ${session.conversationData.company}. Retrieving last 10 conversations ...`;
                session.send(initialOutput);

                var conversations = botdisplay.renderSummaryCard(session, builder, results);

                var outputCards = new builder.Message(session)
                    .attachmentLayout(builder.AttachmentLayout.carousel)
                    .attachments(conversations);

                session.send(outputCards);
            }
            else 
            {
                var emptyResponse = `There are no conversations with ${session.conversationData.company} saved in the archive.\n\n`;
                session.send(emptyResponse);
                session.endDialog("Enter OK to return to the home screen.");
            }

        }, function (err)
        {
            console.error(`Could not retrieve stored conversations for ${session.dialogData.inputCompany}:`, err);
        });

    }
]).triggerAction({ matches: /^Retrieve Entry$/i });

// Launch PowerBI dashboard Url
bot.dialog('/viewDashboard',
[
    function(session)
    {
        if (session.message.address.channelId != 'email')
        {
            if (session.conversationData.pinExists == false)
                session.replaceDialog('/auth')
        
            var dashboardURL = process.env.DashboardUrl;
            if (dashboardURL)
                session.send(`The conversation dashboard is available at: ${dashboardURL}`).endDialog();
            else
                session.send(`Oh no! Looks like I wasn't configured with the location of the conversation dashboard.`).endDialog();
        }
    }
]).triggerAction({ matches: /^Dashboard$/i });

// Display conversation details using adaptive card
bot.dialog('/displayConversationCard',
[
    function(session, args, next)
    {
        var allActionsAvailable = [
            {
                "type": "Action.Submit",
                "title": "Discard conversation",
                "data": {"message": "discard"}
            },
            {
                "type": "Action.Submit",
                "title": "Edit conversation",
                "data": {"message": "edit"}
            },
            {
                "type": "Action.Submit",
                "title": "Confirm",
                "data": { "message" : "confirm"}  
            }
        ];

        var submitDisabled = [
            {
                "type": "Action.Submit",
                "title": "Discard conversation",
                "data": {"message": "discard"}
            },
            {
                "type": "Action.Submit",
                "title": "Edit conversation",
                "data": {"message": "edit"}
            }
        ];   
        
        if (!session.message.value)
        {
            var templateCompleted = templateComplete(session.conversationData);
            var outputCard;

            if (templateCompleted)
                outputCard = botdisplay.renderCard(session, builder, args, allActionsAvailable);
            else
                outputCard = botdisplay.renderCard(session, builder, args, submitDisabled);
            session.send(outputCard);
        }
        
        next();
    },
    function (session, results)
    {
        if ((session.message) && (session.message.value))
        {
            var response = session.message.value;
            var templateCompleted = templateComplete(session.conversationData);
            
            if (response.message === "edit")
            {
                session.message.value = null;
                session.replaceDialog('/displayEditableCard', session.conversationData);
            }
            else if ((response.message === "confirm") && (templateCompleted))
            {
                    session.message.value = null;
                    session.replaceDialog('/confirm');  
            }
            else if ((response.message === "confirm") && (!templateCompleted))
            {
                session.message.value = null;
                session.send("Please complete all required sections in the template. Required sections are marked with an asterisk(*).");
                session.replaceDialog('/displayConversationCard', session.conversationData);
            }
            else if (response.message === "discard")
            {
                session.message.value = null;
                session.send("Discarding conversation details... Goodbye");
                session.endConversation();
            }

            else if (response.message === "confirm")
            {
                session.message.value = null;
                session.send("Please reply with your pin to confirm. (found here https://customerauthbot.azurewebsites.net/)");

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
        var prompt;
        var conversationObject = args || new Object();

        // Confirmation and edit instructions
        if (templateComplete(conversationObject))
        {

            if(session.message.address.channelId === 'email') {
                prompt = `Below you will find the details of your conversation.\n\n`;
                prompt += "`---`\n\n";
                prompt += "Reply with **your pin** [found here](https://customerauthbot.azurewebsites.net/) to accept the conversation details below.\n\n";
                prompt += "Reply with **Discard** to discard the conversation.\n\n";
                prompt += "**Edit the details below** and reply if you would like to change any conversation detail below.\n\n";
                prompt += "`---`\n\n";                
            }
            else {
                prompt = `Below you will find the details of your conversation.\n\n`;
                prompt += "`---`\n\n";
                prompt += "Reply with **Confirm** to accept the conversation details below.\n\n";
                prompt += "Reply with **Discard** to discard the conversation.\n\n";
                prompt += "**Edit the details below** and reply if you would like to change any conversation detail below.\n\n";
                prompt += "`---`\n\n"; 
            }
        }
        else
        {
            prompt = `Below you will find the details of your conversation.\n\n`;
            prompt += "Please complete all required sections in the template to continue. You will find some fields have been populated with hints for acceptable values.\n\n";
            prompt += "`---`\n\n";   
            prompt += "Required sections are **marked with an asterisk('*')**.\n\n" 
            prompt += "Reply with **Discard** to discard the conversation.\n\n";
            prompt += "**Edit the required details below** and reply back to continue.\n\n";
            prompt += "`---`\n\n"; 
        }

        // Conversation details
        var outputMessage = botdisplay.renderText(prompt, conversationObject);
        builder.Prompts.text(session, outputMessage);
    },
    function (session, results)
    {
        //if discard
        if (/^[\s]*discard[\s]*/im.test(session.message.text))
        {
            session.send("Discarding conversation.");
            session.endConversation();
        }

        else if (/^[\s]*confirm[\s]*/im.test(session.message.text) && session.message.address.channelId !== 'email')
        {
            session.replaceDialog('/confirm');
        }

        //if it's numbers check the pin
        else if ((/^[\s]*[0-9]+/gm.test(session.message.text)) && (templateComplete(session.conversationData)) && session.message.address.channelId === 'email')
        {
            
            var confirmationPinRegex = /^[\s]*([0-9]+)/g
            var pinMatches = confirmationPinRegex.exec(session.message.text);
            
            if (pinMatches != null)
                session.dialogData.pin = pinMatches[1];
            
            console.log("This is the pin: ", session.dialogData.pin)
            var email = session.userData.alias
            email += '@microsoft.com' 
            console.log("This is the email: ", email)

            var sqlAuthQuery = `SELECT TOP 1 Code FROM [dbo].[app_authcodes]  WHERE Alias='${email}'`;
            dbconnection.execute({
                query: sqlAuthQuery
            }).then (function (results)
            {
                console.log("This is the query result: ", results[0].Code)
                if (results[0].Code == session.dialogData.pin) {
                    session.conversationData.pinExists = true;
                    session.replaceDialog('/confirm');
                }
                else{
                    session.send("Your alias and pin do not match");   
                    session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);
                }
            }, function (err)
            {
                console.error(`Could not retrieve authcode for your alias`, err);
            });   
        }

        //if else/edit 
        else
            session.replaceDialog('/editConversation', session.message.text);
    }
]);

// Display editable conversation card
bot.dialog('/displayEditableCard',
[
    function (session, args, next)
    {
        var inputCard = botdisplay.renderEditableCard(session, builder, args);
     
        // Display editable card if it hasn't been displayed already
        if (!session.message.value)
            session.send(inputCard);
        
        next();
    },
    function(session, results, next)
    {
        if (session.message.value)
        {
             if (session.message.value.action === "discard")
             {
                session.message.value = null;
                session.send("Discarding conversation details... Goodbye");
                session.endConversation();
             }
             else if (session.message.value.action === "submit")
             {
                var selectedProducts="";
                var inputCompany;
                session.conversationData.contact = session.message.value.contact;
                //session.conversationData.company = session.message.value.company;
                session.conversationData.authors = session.message.value.authors;
                session.conversationData.tags = session.message.value.tags;
                session.conversationData.notes = session.message.value.notes;
                session.conversationData.summary = session.message.value.summary;
                session.conversationData.blockers = session.message.value.blockers;
                session.conversationData.projectstage = session.message.value.projectstage;
                session.conversationData.satisfaction = session.message.value.satisfaction;

                if (session.message.value.ProductVM == "true")
                    selectedProducts += "SQL VM,";
                if (session.message.value.ProductDB == "true")
                    selectedProducts += "SQL DB,";
                if (session.message.value.ProductDW == "true")
                    selectedProducts += "SQL DW,";
                if (session.message.value.ProductElasticPool == "true")
                    selectedProducts += "Elastic pool,";
                if (session.message.value.ProductOnPrem == "true")
                    selectedProducts += "On-prem SQL Server";
                if (session.message.value.ProductOther == "true")
                    selectedProducts += "Other";
                selectedProducts = selectedProducts.replace(/,$/g, "");
                session.conversationData.product = selectedProducts;

                inputCompany = session.message.value.company;
                var promptRegex = /[*]*{Found a few matching companies. Select one of the companies below:}[*]*[\S.]*/ig
                inputCompany = inputCompany.replace(promptRegex, '');
                inputCompany = inputCompany.replace(/[__\r\n]+/g,'');
           
                if ((typeof session.conversationData.companyMatches != "undefined") && (!session.conversationData.customerGuid) && (session.conversationData.companyMatches.includes(inputCompany)))
                {
                    var companyIndex = session.conversationData.companyMatches.indexOf(inputCompany);
                    session.conversationData.company = session.conversationData.companyMatches[companyIndex];
                    session.message.value = null;
                    session.beginDialog('/selectCompany');
                }
                else if (!session.conversationData.customerGuid)
                {
                    session.beginDialog('/findCompanyMatches', session.message.value.company);   
                    session.message.value = null;
                }
                else if (session.conversationData.company != session.message.value.company)
                {
                    session.beginDialog('/findCompanyMatches', session.message.value.company);
                    session.message.value = null;
                }
                else
                {
                    session.message.value = null;
                    next();             
                }
             }
             else
             {
                session.message.value = null;
                session.send("Did not understand that response. Please edit, or discard the conversation details using the provided buttons");
                session.replaceDialog('/displayEditableCard');
             }

        }
    },
    function(session, results, next)
    {
        if (session.conversationData.displayMarkdown)
            session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);
        else if (session.message.address.channelId === "emulator")
            session.replaceDialog('/displayConversationCard', session.conversationData);
        else
            session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);   
    }
]);

// Edit existing conversation details
bot.dialog('/editConversation', [
    function (session, args, next)
    {
        var botChannel = session.message.address.channelId;

        if ((!args) && (botChannel == "emulator"))
        {
            session.send("I did not receive any parameters to change. Resending conversation card...");
            session.replaceDialog('/displayConversationCard', session.conversationData);
        }
        else if (!args)
        {
            session.send("I did not receive any parameters to change. Resending conversation card...");
            session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);
        }
        else
            session.replaceDialog('/batchParser', args);
    }
]);

// Select company name from list of suggested matches
bot.dialog('/selectCompany', [
    function (session, args, next)
    {
        if (session.conversationData.company)
            next();
        else 
        {
            session.dialogData.inputCompanies = args;
            session.send(`I found a few companies that match your input company`);
            var prompt = "Please select one company from the choices below:\n\n";
            builder.Prompts.choice(session, prompt, args);
        }
    },
    function (session, results, next)
    {
        if (results.response)
        {
            session.conversationData.company = results.response.entity;
            getCompanyGUID(session, session.conversationData.company);
        }
        else if (session.conversationData.company && (!session.conversationData.customerGuid))
            getCompanyGUID(session, session.conversationData.company);
        else if (session.conversationData.company && session.conversationData.customerGuid)
            session.endDialog();
        else
            session.replaceDialog('/selectCompany', session.dialogData.inputCompanies)
    }    
]);

// Fetch company whose email domain matches input email
bot.dialog('/findCompanyFromEmail', [
    function (session, args, next)
    {
        if (args)
        {
            session.dialogData.inputEmailDomain = args;
            getCompanyFromEmail(session, session.dialogData.inputEmailDomain);
        }
    }
]);

// Fetch companies matching input company in customer database
bot.dialog('/findCompanyMatches', [
    function (session, args, next)
    {
        if (args)
        {
            session.dialogData.inputCompany = args.trim();
            session.dialogData.inputCompany = session.dialogData.inputCompany.replace(/[\r\n]+/g,'');
            next();
        }
        else 
            builder.Prompts.text(session, "What company did you (all) speak with?");
    },
    function (session, results, next)
    {
        if (results.response && (!session.dialogData.inputCompany))
            session.dialogData.inputCompany = results.response;

        session.sendTyping();
        getCompanyMatches(session, session.dialogData.inputCompany);        
    }
]);

// Store customer email domain in database for easy retrieval later
bot.dialog('/storeEmailDomain', [
    function (session, args, next)
    {
        var insertEmailQuery = `INSERT INTO customer_emaildomain VALUES('${session.conversationData.customerGuid}', '${session.conversationData.inputEmailDomain}');`

        dbconnection.execute('feedbackDb',{
        query: insertEmailQuery        
        }).then (function (results)
        {
            console.log("Saved new customer email domain to database:", session.conversationData.inputEmailDomain);
            session.endDialog();                
        }, function (err)
        {
            console.log(`Could not save customer email domain to database:`, err);
            session.endDialog();
        });  
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
            // Set default properties to null
            if (session.conversationData.authors == "{Microsoft alias}")
                session.conversationData.authors = null;
            if (session.conversationData.company == "{company name}")
                session.conversationData.company = null;
            if (session.conversationData.contact == "{customer contact name}")
                session.conversationData.contact = null;
            if (session.conversationData.product == "{SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other}")
                session.conversationData.product = null;
            if (session.conversationData.notes == "{enter note text here}")
                session.conversationData.notes = null;
            if (session.conversationData.summary == "{enter short summary of note here}")
                session.conversationData.summary = null;
            if (session.conversationData.tags == "{Enter a comma-separated list of tags if any}")
                session.conversationData.tags = null;
            if (session.conversationData.blockers == "{Enter a comma-separated list of blockers if any}")
                session.conversationData.blockers = null;
            if (session.conversationData.projectstage == "{Select one of: Pre-POC, POC, Production}")
                session.conversationData.projectstage = null;
            if (session.conversationData.satisfaction == "{Select a customer satisfaction level from 1-4 with 4 being very satisfied and 1 being very dissatisfied. Enter 0 for unknown}")
                session.conversationData["satisfaction"] = null;
            
            if (session.conversationData.saveCustomerEmail)
                session.beginDialog('/storeEmailDomain');
            storeConversation(session, session.conversationData);
            // Feedback.create({
            //     Name: session.userData.alias,
            //     Authors: session.conversationData.authors,
            //     Company: session.conversationData.company,
            //     Contact: session.conversationData.contact,
            //     Product: session.conversationData.product,
            //     Notes: session.conversationData.notes,
            //     Summary: session.conversationData.summary,
            //     Tags: session.conversationData.tags,
            //     Blockers: session.conversationData.blockers,
            //     ProjectStage: session.conversationData.projectstage
            // }).then(feedback => {console.log(feedback.get({plain: true}))});
                //db_connection.close();
            
        }        
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
        builder.Prompts.text(session, "You can enter OK to continue ...");
    },
    function (session, results)
    {
        session.endDialog();
    }
]).triggerAction({
    matches: /^\/help|^help$/ig, 
    onSelectAction: function (session, args, next) {session.beginDialog('/help', args);}
});

//Restart dialog
bot.dialog('/home', [
    function (session, args, next)
    {
        if (session.conversationData.pinExists == false) {
            session.replaceDialog('/auth')
        }        
        else {
            session.replaceDialog('/selectAction');
        }
    }
]).triggerAction({ matches: /^\/home/i }); 

//Reset dialog
bot.dialog('/reset', [
    function (session, args, next)
    {
        session.perUserInConversationData = {};
        session.userData = {};
        session.conversationData = {};
        session.send("Resetting to default state...");
        session.replaceDialog('/');
    }
]).triggerAction({ matches: /^\/reset/i });

// Konami dialog
bot.dialog('/konami', [
    function (session, args, next)
    {
        var message = "I see you've found the hidden feature.\n\n";
        message += `Congratulations ${session.userData.firstName} 💩`;

        session.send(message);
        session.replaceDialog('/home');
    }
]).triggerAction({ matches: /^UUDDLRLRBA/i});

//Bot Analytics

//answer queries 
// bot.dialog('/query', [
 
//     function(session)
//     {
//         builder.Prompts.text(session, 'Query the Database');
//     },
//     function (session, results)
//     {
//         var sqlQuery = results.response;
        
//         dbconnection.execute({
//             query: sqlQuery
//         }).then (function (results)
//         {
//             console.log("results are: ", results);
            
//             session.endDialog('Resutls %s!', results[0].Authors);
  
//         }, function (err)
//         {
//             console.error(`Query didn't work`, err);
//         });   

//     }


// ]).triggerAction({ matches: /^\/query/i });

// Check if text is email
function isEmail(inputText)
{
    var hasSender = (inputText.search(/from:/i) != -1);
    var hasRecipient = (inputText.search(/to:[\s+\w]+/i) != -1);

    //TODO does it need a subject?
    var hasSubject = (inputText.search(/subject:/i) != -1);

    if (hasSender && hasRecipient && hasSubject)
        return true;
    else
        return false;
}

//Check if text is valid conversation template
function containsTemplateKeyword(inputText)
{
    if (inputText.search(/author[(s)]*?:/i) != -1)
        return true;
    else if (inputText.search(/company:/i) != -1)
        return true;
    else if (inputText.search(/contact[(s)]*?:/i) != -1)
        return true;
    else if (inputText.search(/product[(s)]*?:/i) != -1)
        return true;
    else if (inputText.search(/tags?:/i) != -1)
        return true;
    else if (inputText.search(/notes?:/i) != -1)
        return true;
    else if (inputText.search(/summary:/i) != -1)
        return true;
    else
        return false;
}

function isValidTemplate(inputText)
{
    var hasAuthor = (inputText.search(/author[(s)*]*\*?:/i) != -1);
    var hasCompany = (inputText.search(/company\*?:/igm) != -1);
    var hasContact = (inputText.search(/contact[(s)]*\*?:/i) != -1);
    var hasProduct = (inputText.search(/product[(s)]*\*?:/i) != -1);
    var hasNotes = (inputText.search(/notes\*?:/i) != -1);

    if (hasAuthor && hasCompany && hasContact && hasProduct && hasNotes)
        return true 
    else
        return false
}

// Check if all required fields in template have been completed
function templateComplete(inputTemplate)
{
    if ((inputTemplate.authors) && (inputTemplate.company) && (inputTemplate.contact) && (inputTemplate.product) && (inputTemplate.notes))
        return true;
    else
        return false;
}

//Return default conversationTemplate
function getConversationTemplate()
{
    var conversationTemplate = "";

    conversationTemplate+= "**Authors:** {Microsoft alias}, {Microsoft alias} ... \n\n";
    conversationTemplate+= "**Company:** {company name} \n\n";
    conversationTemplate+= "**Contact:** {customer contact name}, {customer contact name} ... \n\n";
    conversationTemplate+= "**Product:** {SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other} \n\n";
    conversationTemplate+= "**Notes:** {enter note text here} \n\n";
    conversationTemplate+= "**Summary(optional):** {enter short summary of note here}\n\n";
    return conversationTemplate
}

function getCompanyFromEmail(session, inputEmailDomain)
{
    if (inputEmailDomain)
    {
       var companyEmailQuery = `SELECT ms_customer_guid FROM customer_emaildomain WHERE email_domain='${inputEmailDomain.toLowerCase()}'`;

       dbconnection.getPlainContext('feedbackDb')
        .step ('getCompanyGuid', {
            query: companyEmailQuery
        })
        .step ('getCompanyName', function (execute, data) {
            var companyGUIDQuery;

            if (data.getCompanyGuid.length)
            {
                var inputGuid = data.getCompanyGuid[0].ms_customer_guid;
                companyGUIDQuery = `SELECT DISTINCT customerName from [dbo].[WeeklyCustomerExperienceReportv13] WHERE MSCustomerGuid='${inputGuid}'`;
            }
            else
                companyGUIDQuery = `SELECT 'Unknown Company'`
            
            execute({
                query: companyGUIDQuery
            })
        })
        .end( function (results) {
            var companyGuids = results.getCompanyGuid;
            var companyNames = results.getCompanyName;

            if (companyGuids.length)
            {
                session.conversationData.customerGuid = companyGuids[0].ms_customer_guid;
                if (companyNames.length)
                    session.conversationData.company = companyNames[0].customername;
                else
                    session.conversationData.company = inputEmailDomain;
                session.endDialog();
            }
            else
            {
                session.conversationData.saveCustomerEmail = true;
                session.conversationData.inputEmailDomain = inputEmailDomain;

                //sends to "what companay did you all speak with"
                var companyName = inputEmailDomain.split('.')[0];
                session.replaceDialog('/findCompanyMatches', companyName);
            }

        })
        .error (function (err) {
            console.log("Could not search database for provided company email:", err);
            session.send("I have encountered an error while searching for the input company. Restarting...");
            session.endDialog();
            session.beginDialog('/home');
        });      
    }
}

function getCompanyMatches(session, inputCompany)
{
    // Find input company in feedback database
    var sqlCompanyQuery = "SELECT DISTINCT customername from [dbo].[WeeklyCustomerExperienceReportv13];";

    dbconnection.execute('feedbackDb',{
        query: sqlCompanyQuery
    }).then (function (results)
    {
        var companies = results
        var searchOptions = 
        {
            shouldSort: true,
            includeMatches: true,
            minMatchCharLength: 1,
            keys: ["customername"],
            threshold: 0.19,
            distance: 10
        }
        var fuse = new Fuse(companies, searchOptions);

        // Search database for input company name and companies with matching names
        var result = fuse.search(inputCompany);
        var matches = result.map(function (entry) {return entry.item.customername;});

        if (session.conversationData.interactiveMode)
        {
            if (matches.length == 0)
            {
                session.send(`Hmmm. I can't appear to find '${inputCompany}' in the customer database. Please refine your search or try another company.`);
                session.replaceDialog('/findCompanyMatches');
            }    
            else if (matches.length > 1)
            {
                session.beginDialog('/selectCompany', matches.slice(0,20));            
            }
            else if (matches.length == 1)
            {
                session.conversationData.company = matches[0];
                session.beginDialog('/selectCompany', matches);
            }
        }
        else if (matches.length == 1)
        {
            session.conversationData.company = matches[0];
            getCompanyGUID(session, session.conversationData.company);
        }
        else
        {
            console.log("matches:", matches);   
            session.conversationData["companyMatches"] = matches.slice(0,10);  
            session.endDialog();       
        }
    }, function (err)
    {
        console.log("Could not search database for provided company name:", err);
        session.send("I have encountered an error while searching for the company with that name. Restarting...");
        session.endDialog();
        session.beginDialog('/home');
    });
}

// Fetch GUID for input company
function getCompanyGUID(session, inputCompany)
{
    //Find customer GUID for input company
    var companyGUIDQuery = `SELECT TOP 1 MSCustomerGuid FROM [dbo].[WeeklyCustomerExperienceReportv13] where CustomerName LIKE '%${inputCompany}%' ORDER BY LEN(Search) DESC;`

    dbconnection.execute('feedbackDb',{
        query: companyGUIDQuery
    }).then (function (results)
    {
        var companyGUID = results[0].MSCustomerGuid;
        session.conversationData.customerGuid = companyGUID;
        session.endDialog();
    }, function (err)
    {
        console.log(`Could not fetch customer GUID for '${inputCompany}':`, err);
        session.send("I have encountered an error while searching for the ID for the provided customer name. Restarting...");
        session.endDialog();
        session.beginDialog('/home');
    });
}

//Parse conversation template into session variable
function parseConversationTemplate(session, inputText)
{
    console.log('parsing the conversation template');
    var emailSignatureRegex = /(^[\s]*--*[\s]*[a-z \.]*\w+$|^[\s]*best[\s,!\w]*\w+$|^[\s]*regards[\s,!\w]*\w+$|^[\s]*thanks[\s,!\w]*\w+$|^[\s]*cheers[\s,!\w]*\w+$|^[\s]*sent from [\w' ]+$)/im
    var conversationTemplateRegex = /(^author[(s)*]*?:|company[*]?:|contact[(s)*]*?:|customer contact[(s)*]*?:|product[(s)*]*?:|tags?:|tags?[(optional)]+:|notes[*]?:|summary:|summary[(optional)]+:|blocker[(s)]*?:|projectstage:|project stage:|satisfaction:|customer satisfaction:)/im;

    // Parse email signatures out of input text
    var templateTokens = inputText.replace(emailSignatureRegex, '');

    // Parse lines out of input text
    templateTokens = templateTokens.replace(/__/g,'');
    
    //Split input text into tokens
    templateTokens = templateTokens.split(conversationTemplateRegex);
    console.log('tokens:',templateTokens);
    var endToken = templateTokens.length -1;
    var inputCompany;

    for (var token=0; token<templateTokens.length; token++)
    {
        if (templateTokens[token].search(/author[(s)*]*?:/i) != -1)
        {
            // Ignore input if it's default text, include otherwise
            if ((token != endToken) && (templateTokens[token+1].search(/{Microsoft alias}/i) == -1))
            {
                if (templateTokens[token+1].includes(session.userData.alias))
                {
                    var inputAuthors = templateTokens[token+1];
                    inputAuthors = inputAuthors.replace(/[__\r\n]+/g,'');
                    session.conversationData["authors"] = inputAuthors;
                }
                else
                {
                    // Include author alias if it is not included in authors list
                    var authorsList = session.userData.alias + ", " + templateTokens[token+1];
                    authorsList = authorsList.replace(/,$/g, '');
                    authorsList = authorsList.replace(/[__\r\n]+/g,'');
                    session.conversationData["authors"] = authorsList;
                }    
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/company[*]?:/i) != -1))
        {
            // Parse out prompt text
            var promptRegex = /[*]*{Found a few matching companies. Select one of the companies below:}[*]*[\S.]*/ig
            if (templateTokens[token+1].search(promptRegex) != -1)
            {
                inputCompany = templateTokens[token+1];
                inputCompany = inputCompany.replace(promptRegex, '');
                inputCompany = inputCompany.replace(/[__\r\n]+/g,'');
            }
            // Ignore input if it's default text
            else if (templateTokens[token+1].search(/{company name}/i) == -1)
            {
                inputCompany = templateTokens[token+1];
                inputCompany = inputCompany.replace(/[__\r\n]+/g,'');
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/contact[(s)*]*?:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{customer contact name}/i) == -1)
            {
                var inputContacts = templateTokens[token+1];
                inputContacts = inputContacts.replace(/[__\r\n]+/g,'');
                session.conversationData["contact"] = inputContacts;
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/product[(s)*]*?:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other}/i) == -1)
            {
                var inputProducts = templateTokens[token+1];
                inputProducts = inputProducts.replace(/[__\r\n]+/g,'');
                session.conversationData["product"] = inputProducts;
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/tags?:|tags?[(optional)]+:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{Enter a comma-separated list of tags if any}/i) == -1)
            {
                var inputTags = templateTokens[token+1];
                inputTags = inputTags.replace(/[__\r\n]+/g,'');
                session.conversationData["tags"] = inputTags;
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/notes[*]?:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{enter note text here}/i) == -1)
            {
                var inputNotes;

                if (isEmail(templateTokens[token+1]))
                    inputNotes = botdisplay.renderEmailConversation(templateTokens[token+1]);
                else
                    inputNotes = templateTokens[token+1];
                session.conversationData["notes"] = inputNotes;
            } 
        }
        else if ((token != endToken) && (templateTokens[token].search(/summary:|summary[(optional)]+:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{enter short summary of note here}/i) == -1)
            {
                var inputSummary = templateTokens[token+1];
                inputSummary = inputSummary.replace(/[__\r\n]+/g,'');
                session.conversationData["summary"] = inputSummary;
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/blocker[(s)]*?:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{Enter a comma-separated list of blockers if any}/i) == -1)
            {
                var inputBlockers = templateTokens[token+1];
                inputBlockers = inputBlockers.replace(/[__\r\n]+/g,'');
                session.conversationData["blockers"] = inputBlockers;
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/projectstage:|project stage:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{Select one of: Pre-POC, POC, Production}/i) == -1)
            {
                var inputStage = templateTokens[token+1];
                inputStage = inputStage.replace(/[__\r\n]+/g,'');
                session.conversationData["projectstage"] = inputStage;
            }
        }
        else if ((token != endToken) && (templateTokens[token].search(/satisfaction:|customer satisfaction:/i) != -1))
        {
            // Ignore input if it's default text
            if (templateTokens[token+1].search(/{Select a customer satisfaction level from 1-4 with 4 being very satisfied and 1 being very dissatisfied. Enter 0 for unknown}/i) == -1)
            {
                var inputSat = templateTokens[token+1];
                inputSat = inputSat.replace(/[__\r\n]+/g,'');
                session.conversationData["satisfaction"] = inputSat;
            }
        } 
   }

    // Get matches for input company if a company hasn't been selected already
    if ((typeof session.conversationData.companyMatches != "undefined") && (!session.conversationData.customerGuid) && (session.conversationData.companyMatches.includes(inputCompany)))
    {
        var companyIndex = session.conversationData.companyMatches.indexOf(inputCompany);
        session.conversationData.company = session.conversationData.companyMatches[companyIndex];
        session.beginDialog('/selectCompany');
    }
    else if (!session.conversationData.customerGuid)
        session.beginDialog('/findCompanyMatches', inputCompany);
    else if (session.conversationData.company != inputCompany)
        session.beginDialog('/findCompanyMatches', inputCompany); 
    else
        session.beginDialog('/selectCompany');   
}

//Write data to database
function storeConversation(session,inputConversation)
{
    // Convert authors list to json format
    var authorsList = [];
    var tokens = inputConversation.authors.split(',');
    for (var i = 0; i < tokens.length; i++)
    {
        authorsList.push({"alias": tokens[i]});
    }
    if (authorsList.length > 0)
        authorsList =  "N'" + JSON.stringify(authorsList) + "'";
    else
        authorsList = null;

    // Convert customer contacts to json format
    var contactsList = [];
    tokens = inputConversation.contact.split(',');
    for (var i = 0; i < tokens.length; i++)
    {
        var fullname = tokens[i].trim();
        var firstname, lastname;

        if (fullname.match(' '))
        {
            firstname = fullname.substr(0, fullname.indexOf(' '));
            lastname = fullname.substr(fullname.indexOf(' ')).trim();
        }
        else
        {
            firstname = fullname;
            lastname = '';
        }

        contactsList.push({"first_name": firstname, "last_name": lastname});
    }
    if (contactsList.length > 0)
        contactsList = "N'" + JSON.stringify(contactsList) + "'";
    else
        contactsList = null;

    // Create service ID integer for products discussed during conversation
    var serviceId = 0;
    if (inputConversation.product.match(/DB/i))
        serviceId|=1 << 0   //Turn on bit at position 0 for SQL DB
    if (inputConversation.product.match(/DW/i))
        serviceId|=1 << 1   //Turn on bit at position 1 for SQL DW
    if (inputConversation.product.match(/VM/i))
        serviceId|=1 << 2   //Turn on bit at position 2 for SQL VM
    if (inputConversation.product.match(/pool/i))
        serviceId|=1 << 3   //Turn on bit at position 3 for Elastic pool
    if (inputConversation.product.match(/prem/i))
        serviceId|=1 << 4   //Turn on bit at position 4 for On-prem SQL Server
    if (inputConversation.product.match(/other/i))
        serviceId|=1 << 5   //Turn on bit at position 5 for Other

    // Convert all tags to json format
    var convTags = [];
    if (inputConversation.blockers)
    {
        tokens = inputConversation.blockers.split(',');

        for (var i=0; i < tokens.length; i++)
        {
            convTags.push({"tag_type": "Blocker", "tag_value": tokens[i]});
        }
    }

    // Convert projectstage tags to json format
    if (inputConversation.projectstage)
    {
        tokens = inputConversation.projectstage.split(',');
        for (var i=0; i < tokens.length; i++)
        {
            convTags.push({"tag_type": "Project Stage", "tag_value": tokens[i]});
        }
    }

    // Convert search tags to json format
    if (inputConversation.tags)
    {
        tokens = inputConversation.tags.split(',');
        for (var i=0; i < tokens.length; i++)
        {
            convTags.push({"tag_type": "Search", "tag_value": tokens[i]});
        }
    }
    if (convTags.length > 0)
        convTags = "N'" + JSON.stringify(convTags) + "'";
    else
        convTags = null;

    // Initialize variables that will be written to conversation table
    var creator_alias, customerGuid, convNotes, convSummary, convSat;
    creator_alias = customerGuid = convNotes = convSummary = convSat= null;

    if (session.userData.alias)
        creator_alias = JSON.stringify(session.userData.alias);
    if (inputConversation.customerGuid)
        customerGuid = JSON.stringify(inputConversation.customerGuid);
    if (inputConversation.satisfaction)
        convSat = JSON.stringify(inputConversation.satisfaction);
    if (inputConversation.notes)
        convNotes = "N'" + inputConversation.notes.replace(/'/g,"''") + "'";
    if (inputConversation.summary)
        convSummary = "N'" + inputConversation.summary.replace(/'/g,"''") + "'";

    var insertConversationQuery = `EXEC dbo.usp_insert_conversation_from_bot @user=${creator_alias}, @authors=${authorsList}, @ms_customer_guid=${customerGuid}, 
        @service_discussed = ${serviceId}, @notes=${convNotes}, @summary=${convSummary}, @customer_contacts=${contactsList}, @satisfaction=${convSat}, @tags=${convTags}`;

    dbconnection.execute('feedbackDb',{
        query: insertConversationQuery        
    }).then (function (results)
    {
        var conversationId = results[0].conversationId;
        customerGuid = customerGuid.replace(/"/g,'');        
        var companyName = encodeURIComponent(inputConversation.company.trim());
        var webLink = `http://azuswdreamdata:3000/#/customer/${customerGuid}/name/${companyName}/conversation/${conversationId}`;
        var dashboardLink = process.env.DashboardUrl;
        var signoffMessages = [`*Alone we can do so little, together we can do so much.* -Helen Keller`, 
        `*In teamwork, silence isn’t golden. It’s deadly.* -Mark Sanborn`,
        `*A single leaf working alone provides no shade.* -Chuck Page`, 
        `*If I have seen further than others, it is by standing upon the shoulders of giants.* -Isaac Newton`,
        `*A successful team is a group of many hands but of one mind* -Bill Bethel`];

        var message = "I have saved your conversation to the archive.\n\n";
        message += `You can [view archived conversations on the conversation dashboard](${dashboardLink}).\n\n`
        message+= "`---`\n\n 💡 ";
        message += signoffMessages[Math.floor(Math.random()*signoffMessages.length)];
        
        session.send(message);
        session.endConversation();                
    }, function (err)
    {
        console.log(`Could not save conversation to database:`, err);
        session.send("I ran into an error while saving your conversation.");
        if (session.message.address.channelId != 'emulator')
            session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);
        else
            session.replaceDialog('/displayConversationCard', session.conversationData);
    });  
}

module.exports = { 
    connector: connector,
    bot: bot
};
