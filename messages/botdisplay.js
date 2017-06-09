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
        conversationObject["contact"]="";
    if (!conversationObject.product)
        conversationObject["product"]="";
    if (!conversationObject.company)
        conversationObject["company"]="";
    if (!conversationObject.authors)
        conversationObject["authors"]="";
    if (!conversationObject.tags)
        conversationObject["tags"]="";
    if (!conversationObject.notes) 
        conversationObject["notes"]="";
    if (!conversationObject.summary)
        conversationObject["summary"]="";

    return conversationObject;
}

// Renders a conversation card using AdaptiveCard framework
function renderCard(session, builder, inputConversation, inputActions=null)
{
	var conversationObject= initializeConversationObject(inputConversation);
        
    var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + conversationObject.contact + " today where you discussed about how " + conversationObject.product + " is used at " + conversationObject.company + "</s><voice gender = \"female\"></voice>"
    var companyName = conversationObject.company.replace(/(?:\r\n|\r|\n)/g, '');
    companyName = companyName.trim();
    var header = "Conversation with " + companyName;

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
    
    var header = "Conversation with " + conversationObject.company;
    var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + conversationObject.contact + " today where you discussed about how " + conversationObject.product + " is used at " + conversationObject.company + "</s><voice gender = \"female\"></voice>";
    var vmSelected = (conversationObject.product.match(/VM/i) != null);
    var dbSelected = (conversationObject.product.match(/DB/i) != null);
    var dwSelected = (conversationObject.product.match(/DW/i) != null);
    var poolSelected = (conversationObject.product.match(/pool/i) != null);
    var onPremSelected = (conversationObject.product.match(/on-prem/i) != null);
    var otherSelected = (conversationObject.product.match(/other/i) != null);

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
                    "text": `Tags:`
                },
                {
                    "type": "Input.Text",
                    "value": `${conversationObject.tags}`,
                    "isRequired": "false",
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
                    "title": "Submit",
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
		var title = `Conversation with ${currConversation.Company}`;
		var subtitle = `Recorded on: ${currConversation.updatedAt}`;
		var cardText;
		var outputCard;
		
		if (currConversation.Summary)
		{
			cardText = "Summary: ";
			if (currConversation.Summary.length > 60)
			 	cardText += currConversation.Summary.substring(0, 61) + "...";
			else
			 	cardText += currConversation.Summary;
		}
		else 
		{
			cardText = "Notes: ";
			if (currConversation.Notes.length > 60)
				cardText += currConversation.Notes.substring(0, 61) + "...";
			else
				cardText += currConversation.Notes;
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
    outputMessage += `**Conversation with ${companyName}**\n\n`;
    outputMessage += `>COMPANY*:\n\n>${conversationObject.company}\n\n---\n`;
    outputMessage += `>AUTHOR(S)*:\n\n>${conversationObject.authors}\n\n---\n`;
    outputMessage += `>CUSTOMER CONTACT(S)*:\n\n>${conversationObject.contact}\n\n---\n`;
    outputMessage += `>PRODUCT(S)*:\n\n>${conversationObject.product}\n\n---\n`;
    outputMessage += `>TAGS:\n\n>${conversationObject.tags}\n\n---\n`;
    outputMessage += `>SUMMARY:\n\n>${conversationObject.summary}\n\n---\n`;
    outputMessage += `>NOTES*:\n\n>${conversationObject.notes}`;

    return outputMessage;
}

// Renders an email chain into markdown format
function renderEmailConversation(inputEmail)
{
	//Split input text by email tags
	var emailTokenRegex = /(from:)|(sent:)|(to: )|(subject:)/ig;
	var inputEmailTokens = inputEmail.split(emailTokenRegex);
	var outputEmail = "";

	//Filter out undefined and empty values
	inputEmailTokens = inputEmailTokens.filter(n => n);

	for (var i=0; i<inputEmailTokens.length; i++)
	{
		if (inputEmailTokens[i] == "From:")
		{
			if (outputEmail.length)
				outputEmail+= "---\n\n";
			else
			{
	            var emailSignatureRegex = /(^[\s]*--*[\s]*[a-z \.]*\w+$|^[\s]*best[\s,!\w]*\w+$|^[\s]*regards[\s,!\w]*\w+$|^[\s]*thanks[\s,!\w]*\w+$|^[\s]*cheers[\s,!\w]*\w+$|^[\s]*sent from [\w' ]+$)/im
	            var forwardMessage = inputEmail.substring(0, inputEmail.indexOf("From:"));

	            // Trim out email signature from initial forward message
	            forwardMessage = forwardMessage.replace(emailSignatureRegex, '');
	            outputEmail = outputEmail + forwardMessage + "---\n\n";
           	}
			
			outputEmail+= ">**From:** ";
			outputEmail+= inputEmailTokens[i+1].trim();
			outputEmail+= "\n\n";
		}
		else if (inputEmailTokens[i] == "Sent:")
		{
			outputEmail+= ">**Sent:** ";
			outputEmail+= inputEmailTokens[i+1].trim();
			outputEmail+= "\n\n";
		}
		else if (inputEmailTokens[i] == "To: ")
		{
			outputEmail+= ">**To:** ";

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
			outputEmail+= ">**Subject:** ";
			var urlRegex = /(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9]\.[^\s]{2,})/ig

			var subjectBody = inputEmailTokens[i+1];
			subjectBody = subjectBody.replace(/\n|\n\n/g,'\n>');
			
			var urlLinks = urlRegex.exec(subjectBody);

			while (urlLinks != null)
			{
				console.log(urlLinks);
				var destUrl = urlLinks[1];
				var markdownLink = `[(link)](${destUrl})`;
				subjectBody = subjectBody.replace(urlLinks[0], markdownLink);
				urlLinks = urlRegex.exec(subjectBody);
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