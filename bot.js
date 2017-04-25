const builder = require('botbuilder');

const connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

const bot = new builder.UniversalBot(connector, [
    (session, args, next) => {
        const botName = 'MAAV-1';
        const description = `I record notes and insights from the conversations Microsoft employees have with customers.  My purpose is to make it easier to disseminate knowledge about Microsoft customers and win as a team.`;

        session.send(`Hi there! I'm ${botName}`);
        session.send(`In a nutshell, here's what I do:\n\n${description}`);

        builder.Prompts.text(session, `What's your name?`);
    },
    (session, results, next) => {
        session.endConversation(`Welcome, ${results.response}`);
    }
]);

module.exports = { 
    connector: connector,
    bot: bot
};