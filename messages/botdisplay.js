function renderCard(session, builder, inputConversation, inputActions=null)
{
	var conversationObject={};
        
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

    var audioSummary = "<s>You had a meeting with <break strength='weak'/> " + conversationObject.contact + " today where you discussed about how " + conversationObject.product + " is used at " + conversationObject.company + "</s><voice gender = \"female\"></voice>"
    var companyName = conversationObject.company.replace(/(?:\r\n|\r|\n)/g, '');
    companyName = companyName.trim();
    var header = "Conversation with " + companyName;

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
    });

    if (Array.isArray(inputActions))
		outputCard.attachments[0].content.actions = inputActions;

	return outputCard;
}

// Returns a carousel display that contains summary cards of input conversations
function renderSummaryCard(session, builder, inputConversations)
{
	var outputCards = [];
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

function renderText(prompt="", inputConversation)
{
	var conversationObject={};
    var outputMessage = prompt;

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


module.exports = { 
    renderCard: renderCard,
    renderText: renderText,
    renderSummaryCard: renderSummaryCard
};