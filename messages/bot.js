"use strict";
var builder = require("botbuilder");
var botbuilder_azure = require("botbuilder-azure");

var Sequelize = require('sequelize');
var Fuse = require('fuse.js');
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
        Notes: Sequelize.STRING(3000),
        Summary: Sequelize.STRING,
        Tags: Sequelize.STRING
    });

    db_connection.sync().then(function()
    {
        console.log("Created database schema from 'feedback' model");
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


// Create connection to database
var config = 
{
    "user": process.env.DB_ADMIN,
    "password": process.env.DB_PASSWORD,
    "server": process.env.DB_SERVER,
    "database": process.env.DB_NAME,

    options: {encrypt: true}
}
dbconnection.setDefaultConfig(config);

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
            session.replaceDialog('/firstRun');
        }
    },
    function (session, results)
    {
        session.userData.alias = results.response;
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
        var userName = session.userData.firstName;
        var message;

        if (session.message.address.channelId === "email")
        {
            // Send markdown version of conversation card for email channel
            session.conversationData.displayMarkdown = true;

            // Parse email chain if bot was forwarded email
            if (isEmail(session.message.text))
            {
                //var message = `Greetings ${userName},\n\n\n\n`;

                //message+= "I see you've sent a previous conversation in the email body. Give me a few minutes to process this information ...";
                //session.send(message);
                session.replaceDialog('/batchParser', session.message.text);
            }
            // Parse conversation template if bot was sent a template as part of greeting
            else if (containsKeyword(session.message.text))
            {
                //var message = `Greetings ${userName},\n\n\n\n`;

                //message+= "I see you've included a few details from your customer conversation in the email body. Give me a few minutes to process this information ...";
                //session.send(message);
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
            session.send("Which product(s) does your customer use?");
            builder.Prompts.text(session, "Available options (use comma to enter multiple products): {SQL VM, SQL DB, SQL DW, Elastic pool, On-prem SQL Server, Other}");
        }
    },
    function (session, results)
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
        productSelection.replace(/,$/g, "");

        session.conversationData.product = productSelection;
        builder.Prompts.text(session, "Please paste in below any notes you took down during the call");
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
        console.log("input:", args);
        // Parse input email chain
        if ((isEmail(session.message.text)) && !(isValidTemplate(session.message.text)))
        {
            // Email parser function goes here
            
            var emailSenderList = [];
            var emailSenderRegex = /from: ([\w ]+)\[mailto:(\w+@\w+.com)]/ig;
            var emailMatches = emailSenderRegex.exec(session.message.text);
            var msftContacts = "";
            var companyName = "";
            var companyContacts = "";

            // Parse out all the email senders into a list
            console.log("In initial email send -without reply validation");
            while (emailMatches != null)
            {
                var author = {name: emailMatches[1], email: emailMatches[2]};
                emailSenderList.push(author);
                emailMatches = emailSenderRegex.exec(session.message.text)
            }

            for (var i=0; i<emailSenderList.length; i++)
            {
                var author = emailSenderList[i];
                
                //Extract MSFT alias
                if (author.email.includes("@microsoft.com"))
                {
                   var authorAlias = author.email.split("@microsoft.com")[0];
                   if (!msftContacts.includes(authorAlias))
                        msftContacts = msftContacts + authorAlias + ","
                }
                //Extract company name and contact info
                else
                {
                    if (!companyContacts.includes(author.name))
                    {
                        companyContacts = companyContacts + author.name + ","
                        companyName = author.email.split("@")[1].replace(".com", "");
                    }
                }
            }

            // Append user's alias if it's not included in list of authors
            if (!msftContacts.includes(session.userData.alias))
                msftContacts = session.userData.alias + ", " + msftContacts;
            
            //Strip out any trailing commas
            msftContacts = msftContacts.replace(/,$/g, "");
            companyContacts = companyContacts.replace(/,$/g, "");

            //Save extracted information into conversation variables
            session.conversationData["authors"] = msftContacts;
            session.conversationData["company"] = companyName;
            session.conversationData["contact"] = companyContacts;
            session.conversationData["notes"] = botdisplay.renderEmailConversation(session.message.text);
        }
        // Parse response to email conversation template
        else if ((isEmail(session.message.text)) && (isValidTemplate(session.message.text)))
        {
            var emailSignatureRegex = /(^[\s]*--*[\s]*[a-z \.]*\w+$|^[\s]*best[\s,!\w]*\w+$|^[\s]*regards[\s,!\w]*\w+$|^[\s]*thanks[\s,!\w]*\w+$|^[\s]*cheers[\s,!\w]*\w+$|^[\s]*sent from [\w' ]+$)/im
            var conversationTemplateRegex = /(author[(s)*]*?:|company[*]?:|contact[(s)*]*?:|customer contact[(s)*]*?:|product[(s)*]*?:|tags?:|tags?[(optional)]+:|notes[*]?:|summary:|summary[(optional)]+:)/i;

            // Parse email signatures out of input text
            console.log("In email reply template parser");
            var templateTokens = session.message.text.replace(emailSignatureRegex, '');
            templateTokens = templateTokens.split(conversationTemplateRegex);
            console.log(templateTokens);

            for (var token=0; token<templateTokens.length; token++)
            {
                if (templateTokens[token].search(/author[(s)*]*?:/i) != -1)
                {
                    // Ignore input if it's default text, include otherwise
                    if (templateTokens[token+1].search(/{Microsoft alias}/i) == -1)
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
                else if (templateTokens[token].search(/company[*]?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{company name}/i) == -1)
                    {
                        var inputCompany = templateTokens[token+1];
                        inputCompany = inputCompany.replace(/[__\r\n]+/g,'');
                        session.conversationData["company"] = inputCompany;
                    }
                }
                else if (templateTokens[token].search(/contact[(s)*]*?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{customer contact name}/i) == -1)
                    {
                        var inputContacts = templateTokens[token+1];
                        inputContacts = inputContacts.replace(/[__\r\n]+/g,'');
                        session.conversationData["contact"] = inputContacts;
                    }
                }
                else if (templateTokens[token].search(/product[(s)*]*?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other}/i) == -1)
                    {
                        var inputProducts = templateTokens[token+1];
                        inputProducts = inputProducts.replace(/[__\r\n]+/g,'');
                        session.conversationData["product"] = inputProducts;
                    }
                }
                else if (templateTokens[token].search(/tags?:|tags?[(optional)]+:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{tag}/i) == -1)
                    {
                        var inputTags = templateTokens[token+1];
                        inputTags = inputTags.replace(/[__\r\n]+/g,'');
                        session.conversationData["tags"] = inputTags;
                    }
                }
                else if (templateTokens[token].search(/notes[*]?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{enter note text here}/i) == -1)
                    {
                        var inputNotes = templateTokens[token+1];
                        inputNotes = inputNotes.replace(/^[\r\n]+/g,'');
                        session.conversationData["notes"] = inputNotes;
                    }    
                }
                else if (templateTokens[token].search(/summary:|summary[(optional)]+:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{enter short summary of note here}/i) == -1)
                    {
                        var inputSummary = templateTokens[token+1];
                        inputSummary = inputSummary.replace(/[__\r\n]+/g,'');
                        session.conversationData["summary"] = inputSummary;
                    }
                }
            }

        }
        // Parse input conversation template
        else
        {
            console.log("In parse input conversation mode");
            session.sendTyping();
            var emailSignatureRegex = /(^[\s]*--*[\s]*[a-z \.]*\w+$|^[\s]*best[\s,!\w]*\w+$|^[\s]*regards[\s,!\w]*\w+$|^[\s]*thanks[\s,!\w]*\w+$|^[\s]*cheers[\s,!\w]*\w+$|^[\s]*sent from [\w' ]+$)/im
            var conversationTemplateRegex = /(\w+[()\w]*:\s*)/i;

            // Parse email signatures out of input text
            var templateTokens = args.replace(emailSignatureRegex, '');
        
            // Parse lines out of input text
            templateTokens = templateTokens.replace(/__/g,'');
                   
            templateTokens = templateTokens.split(conversationTemplateRegex);
            
            for (var token = 0; token<templateTokens.length; token++)
            {
                if (templateTokens[token].search(/author[(s)*]*?:/i) != -1)
                {
                    // Ignore input if it's default text, include otherwise
                    if (templateTokens[token+1].search(/{Microsoft alias}/i) == -1)
                    {
                        if (templateTokens[token+1].includes(session.userData.alias))
                            session.conversationData["authors"] = templateTokens[token+1];
                        else
                        {
                            // Include author alias if it is not included in authors list
                            var authorsList = session.userData.alias + "," + templateTokens[token+1];
                            authorsList = authorsList.replace(/,$/g, "");
                            session.conversationData["authors"] = authorsList;
                        }    
                    }
                }
                else if (templateTokens[token].search(/company[*]?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{company name}/i) == -1)
                        session.conversationData["company"] = templateTokens[token+1];
                }
                else if (templateTokens[token].search(/contact[(s)*]*?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{customer contact name}/i) == -1)
                        session.conversationData["contact"] = templateTokens[token+1];
                }
                else if (templateTokens[token].search(/product[(s)*]*?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other}/i) == -1)
                        session.conversationData["product"] = templateTokens[token+1];
                }
                else if (templateTokens[token].search(/tags?:|tags?[(optional)]+:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{tag}/i) == -1)
                        session.conversationData["tags"] = templateTokens[token+1];
                }
                else if (templateTokens[token].search(/notes[*]?:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{enter note text here}/i) == -1)
                        session.conversationData["notes"] = templateTokens[token+1];
                }
                else if (templateTokens[token].search(/summary:|summary[(optional)]+:/i) != -1)
                {
                    // Ignore input if it's default text
                    if (templateTokens[token+1].search(/{enter short summary of note here}/i) == -1)
                        session.conversationData["summary"] = templateTokens[token+1];
                }
            }
        }

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
        if (session.message.text.match(/^conv#/))
        {
            var selection = session.message.text.split(':')[1].trim();
            var selectedConversation = session.conversationData.retrievedConversations[selection];
            var outputCard;

            if (session.message.address.channelId === "emulator")
                outputCard = botdisplay.renderCard(session, builder, selectedConversation);
            else
                outputCard = botdisplay.renderText("", selectedConversation);

            session.send(outputCard);
            session.conversationData.retrievedConversations = null;
            session.endDialog("Enter OK to return to home screen");
        }
        else if (process.env.DB_SERVER)
            builder.Prompts.text(session, "Which company would you like to retrieve conversations for?");
        else if (!process.env.DB_SERVER)
            session.endDialog("This function is not available. I have not been hooked up to a database of previous conversations.");
    },
    function (session, results, next)
    {
        // Get input company to search for
        if(!session.dialogData.inputCompany)
            session.dialogData.inputCompany = results.response;

        // Find input company in feedback database
        var sqlCompanyQuery = "SELECT DISTINCT company FROM feedbacks;";

        dbconnection.execute({
            query: sqlCompanyQuery
        }).then (function (results)
        {
            var companies = results
            var searchOptions = 
            {
                shouldSort: true,
                includeMatches: true,
                minMatchCharLength: 1,
                keys: ["company"]
            }
            var fuse = new Fuse(companies, searchOptions);

            // Search database for input company name and companies with matching names
            var result = fuse.search(session.dialogData.inputCompany);
            var matches = result.map(function (entry) {return entry.item.company;});
            console.log(matches);

            if (matches.length == 0)
            {
                session.send(`There is no record of '${session.dialogData.inputCompany}' in the database. Please refine your search or try another company.`);
                session.endDialog("Enter OK to return to home screen and try again.");
            }    
            else if (matches.length > 1)
            {
                session.send(`I found a few companies that match '${session.dialogData.inputCompany}`);
                var prompt = "Please select one company from the choices below:\n\n";
                builder.Prompts.choice(session, prompt, matches);
            }
            else if (matches.length == 1)
            {
                var prompt = `Found matching company '${matches}'. Retrieving last 10 conversations...`;
                session.send(prompt);
                session.dialogData.inputCompany = matches;
                next();
            }
        }, function (err)
        {
            console.error("Could not search database for provided company name:", err);
        });
    },
    function (session, results)
    {
        if (results.response)
            session.dialogData.inputCompany = results.response.entity;

        var conversationListQuery = `SELECT TOP 10 [id], [Name],[Authors],[Company],[Contact],[Product],[Notes],[Summary],[Tags],[updatedAt] 
                                     FROM [dbo].[feedbacks] 
                                     WHERE [Company] = '${session.dialogData.inputCompany}'
                                     ORDER BY [updatedAt] DESC`;

        dbconnection.execute({
            query: conversationListQuery
        }).then (function (results)
        {
            var conversations = botdisplay.renderSummaryCard(session, builder, results);

            var outputCards = new builder.Message(session)
                .attachmentLayout(builder.AttachmentLayout.carousel)
                .attachments(conversations);

            session.send(outputCards);

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
        var dashboardURL = process.env.DashboardUrl;
        if (dashboardURL)
            session.send(`The conversation dashboard is available at: ${dashboardURL}`).endDialog();
        else
            session.send(`Oh no! Looks like I wasn't configured with the location of the conversation dashboard.`).endDialog();
    }
]).triggerAction({ matches: /^Dashboard$/i });

// Display conversation details using adaptive card
bot.dialog('/displayConversationCard',
[
    function(session, args, next)
    {
        var availableActions = [
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
        ]
        var outputCard = botdisplay.renderCard(session, builder, args, availableActions)        
        
        if (!session.message.value)
        {
            //session.send("Please confirm or edit the conversation details below:");
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
            prompt = `Below you will find the details of your conversation.\n\n`;
            prompt += "`---`\n\n";
            prompt += "Reply with **Confirm** to accept the conversation details below.\n\n";
            prompt += "Reply with **Discard** to discard the conversation.\n\n";
            prompt += "**Edit the details below** and reply if you would like to change any conversation detail below.\n\n\n\n";
            prompt += "`---`\n\n"; 
        }
        else
        {
            prompt = `Below you will find the details of your conversation. Some fields have been populated with default values\n\n`;
            prompt += "Please complete all required sections in the template to continue.\n\n";
            prompt += "`---`\n\n";   
            prompt += "Required sections are marked with an asterisk(\*).\n" 
            prompt += "Reply with **Discard** to discard the conversation.\n";
            prompt += "**Edit the required details below** and reply back to continue.\n\n\n\n";
            prompt += "`---`\n\n"; 
        }

        // Conversation details
        var outputMessage = botdisplay.renderText(prompt, conversationObject);
        builder.Prompts.text(session, outputMessage);
    },
    function (session, results)
    {
        if ((/^[\s]*confirm[\s]*/im.test(session.message.text)) && (templateComplete(session.conversationData)))
            session.replaceDialog('/confirm');
        else if (/^[\s]*discard[\s]*/im.test(session.message.text))
        {
            session.send("Discarding conversation.");
            session.endConversation();
        }
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
    function(session, results)
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
                session.conversationData.contact = session.message.value.contact;
                session.conversationData.company = session.message.value.company;
                session.conversationData.authors = session.message.value.authors;
                session.conversationData.tags = session.message.value.tags;
                session.conversationData.notes = session.message.value.notes;
                session.conversationData.summary = session.message.value.summary;

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

                session.message.value = null;

                if (session.conversationData.displayMarkdown)
                    session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);
                else if (session.message.address.channelId === "emulator")
                    session.replaceDialog('/displayConversationCard', session.conversationData);
                else
                    session.replaceDialog('/displayMarkdownConversationCard', session.conversationData);
             }
             else
             {
                session.message.value = null;
                session.send("Did not understand that response. Please edit, or discard the conversation details using the provided buttons");
                session.replaceDialog('/displayEditableCard');
             }

        }
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
        session.replaceDialog('/selectAction');
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

// Check if text is email
function isEmail(inputText)
{
    var hasSender = (inputText.search(/from:/i) != -1);
    var hasSendDate = (inputText.search(/sent:/i) != -1);
    var hasRecipient = (inputText.search(/to:[\s+\w]+/i) != -1);
    var hasSubject = (inputText.search(/subject:/i) != -1);

    if (hasSender && hasSendDate && hasRecipient && hasSubject)
        return true;
    else
        return false;
}

//Check if text is valid conversation template
function containsKeyword(inputText)
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
    var hasAuthor = (inputText.search(/author[(s)]*\*?:/i) != -1);
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
    conversationTemplate+= "**Tags(optional):** {tag}, {tag} ... \n\n";
    conversationTemplate+= "**Summary(optional):** {enter short summary of note here}\n\n";
    conversationTemplate+= "**Notes:** {enter note text here} \n\n";
    return conversationTemplate
}



module.exports = { 
    connector: connector,
    bot: bot
};
