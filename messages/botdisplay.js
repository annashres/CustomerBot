// Returns a conversation object that is used to render cards and markdown text
function initializeConversationObject(inputConversation)
{
	var conversationObject={};
        
    //Initialize conversation object to have property names be lowercase
    if (inputConversation)
    {
    	var key, keys = Object.keys(inputConversation);
		var properties = keys.length;
		
		while (properties--)
		{
  			key = keys[properties];
  			conversationObject[key.toLowerCase()] = inputConversation[key];
  		}
    }	

    // Initialize empty property values
    if (!conversationObject.contact) 
        conversationObject["contact"]="**{Enter a customer contact name}**";
    if (!conversationObject.product)
        conversationObject["product"]="**{Select from: SQL VM, SQL DB, SQL DW, Elastic pool, On-Prem SQL Server, Other}**";
    if (!conversationObject.authors)
        conversationObject["authors"]="**{Enter a Microsoft alias}**";
    if (!conversationObject.tags)
        conversationObject["tags"]="{Enter a comma-separated list of tags if any}";
    if (!conversationObject.notes) 
        conversationObject["notes"]="**{Enter note text here}**";
    if (!conversationObject.summary)
        conversationObject["summary"]="{Enter short summary of note here}";
    if (!conversationObject.projectstage)
    	conversationObject["projectstage"]="{Select one of: Pre-POC, POC, Production}";
    if (!conversationObject.blockers)
    	conversationObject["blockers"]="{Enter a comma-separated list of blockers if any}";
    if (!conversationObject.satisfaction)
    	conversationObject["satisfaction"]="{Select a customer satisfaction level from 1-4 with 4 being very satisfied and 1 being very dissatisfied. Enter 0 for unknown}";
    if ((!conversationObject.company) && (!conversationObject.companymatches))
        conversationObject["company"]="**{Enter company name here}**";
    else if ((!conversationObject.company) && conversationObject.companymatches)
    {
        if (conversationObject.companymatches.length > 0)
        {
            var companyPrompt = "**{Found a few matching companies. Select one of the companies below:}**\n\n";

            for (var i = 0; i < conversationObject.companymatches.length; i++)
            {
                companyPrompt =  companyPrompt + conversationObject.companymatches[i] + "\n\n";
            }
        }
        else
            var companyPrompt = "**{No existing companies match the input name. Enter a valid company name}**";
        
        conversationObject["company"]= companyPrompt;
    }    

    return conversationObject;
}

// Renders a conversation card using AdaptiveCard framework
function renderCard(session, builder, inputConversation, inputActions=null)
{
	var conversationObject= initializeConversationObject(inputConversation);
        
    var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + conversationObject.contact + " today where you discussed about how " + conversationObject.product + " is used at " + conversationObject.company + "</s><voice gender = \"female\"></voice>"
    var companyName = conversationObject.company.replace(/(?:\r\n|\r|\n)/g, '');
    companyName = companyName.trim();
    if (companyName.includes("{"))
        companyName = "Unknown company";
    var header = "Conversation with " + companyName + "\n\n";

    if (typeof inputConversation.updatedAt != "undefined")
    {
    	if (inputConversation.updatedAt)
        {
            var recordedDate = new Date(Date.parse(inputConversation.updatedAt)).toString();
      		header += `(Recorded on: ${recordedDate})\n\n`;
        }
    }

    var cardContent = {
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
                    "text": `**Company:\***\n\n ${conversationObject.company}`,
                    "wrap": "true"
                },
                {
                    "type": "TextBlock",
                    "text": `**Author(s):\***\n\n ${conversationObject.authors}`,
                    "separation": "strong",
                    "wrap": "true"
                },
                {
                    "type": "TextBlock",
                    "text": `**Customer contact(s):\***\n\n ${conversationObject.contact}`,
                    "separation": "strong",
                    "wrap": "true"
                },
                {
                    "type": "TextBlock",
                    "text": `**Product(s):\***\n\n ${conversationObject.product}`,
                    "separation": "strong",
                    "wrap": "true"
                },
                {
                	"type": "TextBlock",
                	"text": `Project stage:\n\n ${conversationObject.projectstage}`,
                	"separation": "strong",
                	"wrap": "true"
                },
                {
                	"type": "TextBlock",
                	"text": `Blocker(s):\n\n ${conversationObject.blockers}`,
                	"separation": "strong",
                	"wrap": "true"
                },
                {
                	"type": "TextBlock",
                	"text": `Customer satisfaction:\n\n ${conversationObject.satisfaction}`,
                	"separation": "strong",
                	"wrap": "true"
                },
                {
                    "type": "TextBlock",
                    "text": `Tags:\n\n ${conversationObject.tags}`,
                    "separation": "strong",
                    "wrap": "true"
                },         
                {
                    "type": "TextBlock",
                    "text": `Summary:\n\n ${conversationObject.summary}`,
                    "separation": "strong",
                    "wrap": "true"
                },
                {
                    "type": "TextBlock",
                    "text": `**Notes:\***\n\n ${conversationObject.notes}`,
                    "separation": "strong",
                    "wrap": "true"
                }
            ]
                
        }
    }
    if (Array.isArray(inputActions))
		cardContent.content.actions = inputActions;

    var outputCard = new builder.Message(session)
        .addAttachment(cardContent);

	return outputCard;
}

// Renders an editable conversation card that allows user to edit card contents
function renderEditableCard(session, builder, inputConversation)
{
	var conversationObject= initializeConversationObject(inputConversation);
    
    var companyName = conversationObject.company.replace(/(?:\r\n|\r|\n)/g, '');
    companyName = companyName.trim();
    if (companyName.includes("{"))
        companyName = "Unknown company";
    
    var header = "Conversation with " + companyName;
    var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + conversationObject.contact + " today where you discussed about how " + conversationObject.product + " is used at " + conversationObject.company + "</s><voice gender = \"female\"></voice>";
    var vmSelected, dbSelected, dwSelected, poolSelected, onPremSelected, otherSelected;
    var productPrompt = (conversationObject.product.match(/[{}]/i) != null);

    if (productPrompt)
    {
        vmSelected = dbSelected = dwSelected = poolSelected = onPremSelected = otherSelected = false;
    }
    else
    {
        vmSelected = (conversationObject.product.match(/VM/i) != null);
        dbSelected = (conversationObject.product.match(/DB/i) != null);
        dwSelected = (conversationObject.product.match(/DW/i) != null);
        poolSelected = (conversationObject.product.match(/pool/i) != null);
        onPremSelected = (conversationObject.product.match(/on-prem/i) != null);
        otherSelected = (conversationObject.product.match(/other/i) != null);
    }
    
    var prePOCSelected = (conversationObject.projectstage.match(/pre-poc/i) != null);
    var POCSelected = (conversationObject.projectstage.match(/poc/i) != null);
    var productionSelected = (conversationObject.projectstage.match(/production/i) != null);

    var VSSelected = (conversationObject.satisfaction.match(/4/i) != null);
    var SSSelected = (conversationObject.satisfaction.match(/3/i) != null);
    var SDSelected = (conversationObject.satisfaction.match(/2/i) != null);
    var VDSelected = (conversationObject.satisfaction.match(/1/i) != null);
    var USelected = (conversationObject.satisfaction.match(/0/i) != null);


    var inputCard = new builder.Message(session)
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
                    "text": `**Company:\***`
                },
                {
                    "type": "Input.Text",
                    "value": `${conversationObject.company}`,
                    "isRequired": "true",
                    "id": "company"
                },
                {
                    "type": "TextBlock",
                    "text": `**Author(s):\***`
                },
                {
                    "type": "Input.Text",
                    "value": `${conversationObject.authors}`,
                    "isRequired": "true",
                    "isMultiline": "true",
                    "id": "authors"
                },
                {
                    "type": "TextBlock",
                    "text": `**Customer contact(s):\***`
                },
                {
                    "type": "Input.Text",
                    "value": `${conversationObject.contact}`,
                    "isRequired": "true",
                    "id": "contact"
                },
                {
                    "type": "TextBlock",
                    "text": `**Product(s):\***`
                },
                {
                    "type": "Input.Toggle",
                    "id": "ProductVM",
                    "title": "SQL VM",
                    "value": `${vmSelected}`,
                    "valueOn": "true",
                    "valueOff": "false"
                },
                {
                    "type": "Input.Toggle",
                    "id": "ProductDB",
                    "title": "SQL DB",
                    "value": `${dbSelected}`,
                    "valueOn": "true",
                    "valueOff": "false"
                },
                {
                    "type": "Input.Toggle",
                    "id": "ProductDW",
                    "title": "SQL DW",
                    "value": `${dwSelected}`,
                    "valueOn": "true",
                    "valueOff": "false"
                },
                {
                    "type": "Input.Toggle",
                    "id": "ProductElasticPool",
                    "title": "Elastic pool",
                    "value": `${poolSelected}`,
                    "valueOn": "true",
                    "valueOff": "false"
                },
                {
                    "type": "Input.Toggle",
                    "id": "ProductOnPrem",
                    "title": "On-prem SQL Server",
                    "value": `${onPremSelected}`,
                    "valueOn": "true",
                    "valueOff": "false"
                },
                {
                    "type": "Input.Toggle",
                    "id": "ProductOther",
                    "title": "Other",
                    "value": `${otherSelected}`,
                    "valueOn": "true",
                    "valueOff": "false"
                },
                {
                    "type": "TextBlock",
                    "text": `Project Stage:`
                },
                {
      				"type": "Input.ChoiceSet",
      				"id": "projectstage",
      				"style": "compact",
      				"isMultiSelect": false,
      				"value": `${conversationObject.projectstage}`,
      				"choices":
      				[
				        {
				          "title": "Pre-ProofOfConcept",
				          "value": "Pre-POC",
				          "isSelected": prePOCSelected
				        },
				        {
				          "title": "ProofOfConcept",
				          "value": "POC",
				          "isSelected": POCSelected
				        },
				        {
				          "title": "Production",
				          "value": "Production",
				          "isSelected": productionSelected
				        }
			      	]
    			},
                {
                    "type": "TextBlock",
                    "text": `Customer Satisfaction:`
                },
                {
      				"type": "Input.ChoiceSet",
      				"id": "satisfaction",
      				"style": "compact",
      				"isMultiSelect": false,
      				"value": `${conversationObject.satisfaction}`,
      				"choices":
      				[
				        {
				          "title": "4 - Very Satisfied",
				          "value": "4",
				          "isSelected": VSSelected
				        },
				        {
				          "title": "3 - Somewhat Satisfied",
				          "value": "4",
				          "isSelected": SSSelected
				        },
				        {
				          "title": "2 - Somewhat Dissatisfied",
				          "value": "3",
				          "isSelected": SDSelected
				        },
				        {
				          "title": "1 - Very Dissatisfied",
				          "value": "2",
				          "isSelected": VDSelected
				        },
				        {
				          "title": "0 - Unknown",
				          "value": "0",
				          "isSelected": USelected
				        },
			      	]
    			},            
                {
                	"type": "TextBlock",
                	"text": `Blocker(s):`
                },
                {
                	"type": "Input.Text",
                    "value": `${conversationObject.blockers}`,
                    "isRequired": "false",
                    "isMultiline": "true",
                    "id": "blockers"
                },
                {
                    "type": "TextBlock",
                    "text": `Tags:`
                },
                {
                    "type": "Input.Text",
                    "value": `${conversationObject.tags}`,
                    "isRequired": "false",
                    "isMultiline": "true",
                    "id": "tags"
                },
                {
                    "type": "TextBlock",
                    "text": `Summary:`
                },
                {
                    "type": "Input.Text",
                    "value": `${conversationObject.summary}`,
                    "isRequired": "false",
                    "isMultiline": "true",
                    "id": "summary"
                },
                {
                    "type": "TextBlock",
                    "text": `**Notes:\***`
                },
                {
                    "type": "Input.Text",
                    "value": `${conversationObject.notes}`,
                    "isRequired": "true",
                    "isMultiline": "true",
                    "id": "notes"
                }
            ],
            actions:
            [
                {
                    "type": "Action.Submit",
                    "title": "Discard conversation",
                    "data": {"action": "discard"}
                },
                {
                    "type": "Action.Submit",
                    "title": "Confirm edit",
                    "data": {"action": "submit"}
                }
            ]
        }
    });
	return inputCard;
}

// Returns a carousel display that contains summary cards of input conversations
function renderSummaryCard(session, builder, inputConversations)
{
	var outputCards = [];

    //Store input conversations in session variable for retrieval later
	session.conversationData.retrievedConversations = inputConversations;

	for (var conversation = 0; conversation<inputConversations.length; conversation++)
	{
		var currConversation = inputConversations[conversation];
		var title = `Conversation with ${currConversation.company}`;
		var recordedDate = new Date(Date.parse(currConversation.updatedAt)).toString();
		var subtitle = `Recorded on: ${recordedDate}`;
		var cardText;
		var outputCard;
		
		if (currConversation.summary && (currConversation.summary != 'NULL'))
		{
			cardText = "Summary: ";
			if (currConversation.summary.length > 60)
			 	cardText += currConversation.summary.substring(0, 61) + "...";
			else
			 	cardText += currConversation.summary;
		}
		else 
		{
			cardText = "Notes: ";
            if (!currConversation.notes)
                cardText += "N/A";
			else if (currConversation.notes.length > 60)
				cardText += currConversation.notes.substring(0, 61) + "...";
			else
				cardText += currConversation.notes;
		}

		outputCard = new builder.ThumbnailCard(session)
			.title(title)
			.subtitle(subtitle)
			.text(cardText)
			.buttons([builder.CardAction.imBack(session, `conv#: ${conversation}`, "Select")])

		outputCards.push(outputCard);
	}
	return outputCards
}

// Renders a markdown version of conversation card for channels that don't support AdaptiveCard framework
function renderText(prompt="", inputConversation)
{
	var outputMessage = prompt;
	var conversationObject = initializeConversationObject(inputConversation);

    // Conversation details
    var companyName = conversationObject.company.replace(/(?:\r\n|\r|\n)/g, '');
    companyName = companyName.trim();

    if (companyName.includes("{"))
        companyName = "Unknown company";
    
    var header = "Conversation with " + companyName + "\n\n";
    outputMessage += `**Conversation with ${companyName}**\n\n`;

    // Add recorded date if conversation was previously saved
    if (typeof inputConversation.updatedAt != "undefined")
    {
    	if (inputConversation.updatedAt)
        {
            var recordedDate = new Date(Date.parse(inputConversation.updatedAt)).toString();
      		outputMessage += `(Recorded on: ${recordedDate})\n\n`;
        }
    }

    outputMessage += `COMPANY*:\n\n${conversationObject.company}\n\n---\n`;
    outputMessage += `AUTHOR(S)*:\n\n${conversationObject.authors}\n\n---\n`;
    outputMessage += `CUSTOMER CONTACT(S)*:\n\n${conversationObject.contact}\n\n---\n`;
    outputMessage += `PRODUCT(S)*:\n\n${conversationObject.product}\n\n---\n`;
    outputMessage += `PROJECT STAGE:\n\n${conversationObject.projectstage}\n\n---\n`;
    outputMessage += `BLOCKER(S):\n\n${conversationObject.blockers}\n\n---\n`;
    outputMessage += `SATISFACTION:\n\n${conversationObject.satisfaction}\n\n---\n`;
    outputMessage += `TAGS:\n\n${conversationObject.tags}\n\n---\n`;
    outputMessage += `SUMMARY:\n\n${conversationObject.summary}\n\n---\n`;
    outputMessage += `NOTES*:\n\n${conversationObject.notes}`;

    if (conversationObject.notes.search(/konamicode/ig) != -1)
    {
    	var botSignature = "<div>&#128169</div>";
    	outputMessage = outputMessage + `\n\nThanks ${session.userData.name} ` + botSignature;
    }
    return outputMessage;
}

// Renders an email chain into markdown format
function renderEmailConversation(inputEmail)
{
	//Remove line delimiters from input
	var inputEmailTokens = inputEmail.replace(/__/g,'');
	inputEmailTokens = inputEmailTokens.replace(/----/g,'');

	//Split input text by email tags
	var emailTokenRegex = /(from:)|(sent:)|(to: )|(subject:)/ig;
	inputEmailTokens = inputEmailTokens.split(emailTokenRegex);
	var outputEmail = "";

	//Filter out undefined and empty values
	inputEmailTokens = inputEmailTokens.filter(n => n);

	for (var i=0; i<inputEmailTokens.length; i++)
	{
		if (inputEmailTokens[i] == "From:")
		{
			if (outputEmail.length)
				outputEmail+= "\n\n`----`\n\n";
			else
			{
	            var emailSignatureRegex = /(^[\s]*--*[\s]*[a-z \.]*\w+$|^[\s]*best[\s,!\w]*\w+$|^[\s]*regards[\s,!\w]*\w+$|^[\s]*thanks[\s,!\w]*\w+$|^[\s]*cheers[\s,!\w]*\w+$|^[\s]*sent from [\w' ]+$)/im
	            var forwardMessage = inputEmail.substring(0, inputEmail.indexOf("From:"));
	            forwardMessage = forwardMessage.replace(/----/g,'');

	            // Trim out email signature from initial forward message
	            forwardMessage = forwardMessage.replace(emailSignatureRegex, '');
	            outputEmail = forwardMessage + "\n\n`----`\n\n";;
           	}
			
			outputEmail+= "From: ";
			outputEmail+= inputEmailTokens[i+1].trim();
			outputEmail+= "\n\n";
		}
		else if (inputEmailTokens[i] == "Sent:")
		{
			//Discard this information (needed to deal with Outlook default formatting for emails)
		}
		else if (inputEmailTokens[i] == "To: ")
		{
			outputEmail+= "To: ";

			var emailMisformatRegex = /<mailto:[<\w:@.]+>/gm;
			var respondents = inputEmailTokens[i+1].trim();
			var misformattedEmail = emailMisformatRegex.exec(respondents);

			// Correct email misformat errors that occurred during initial parsing
			respondents = respondents.replace(emailMisformatRegex,'');
		
			outputEmail+= respondents;
			outputEmail+= "\n\n";
		}
		else if (inputEmailTokens[i] == "Subject:")
		{
			outputEmail+= "Subject: ";
			var urlRegex = /<(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9]\.[^\s]{2,})>/ig

			var subjectBody = inputEmailTokens[i+1];
			var urlLinks = urlRegex.exec(subjectBody);

			while (urlLinks != null)
			{
				var destUrl = urlLinks[1];
				var markdownLink = `[(link)](${destUrl})`;
				var existingLinkText = '(link)'+urlLinks[0];

				if (subjectBody.includes(existingLinkText))
					subjectBody = subjectBody.replace(existingLinkText, markdownLink);
				else
					subjectBody = subjectBody.replace(urlLinks[0], markdownLink);
				urlLinks = urlRegex.exec(subjectBody);
			}
			var newLineRegex = /^[a-zA-Z'.,-: ()\?]+(\r\n)/gm;
			var newLineMatch = newLineRegex.exec(subjectBody);

			while (newLineMatch != null)
			{
				subjectBody = subjectBody.replace(newLineMatch[1],'\r\n\r\n');
				newLineMatch = newLineRegex.exec(subjectBody);
			}

			outputEmail+= subjectBody;
			outputEmail+= "\n\n";
		}
	}
	return outputEmail;
}


module.exports = { 
    renderCard: renderCard,
    renderText: renderText,
    renderSummaryCard: renderSummaryCard,
    renderEditableCard: renderEditableCard,
    renderEmailConversation: renderEmailConversation
};