// This module sets up a local API server to run the maav-1 bot on a local dev machine, or binds the bot to a function app if deployed to Azure

// Grab environment variables from .env file
require('dotenv').config();
var useEmulator = (process.env.NODE_ENV == 'development');

// Grab bot file 
var bot = require('./bot.js');

// Create a local API server if the environment is local dev machine
if (useEmulator)
{
	var restify = require('restify');
	var server = restify.createServer();

	server.listen(process.env.PORT, function()
	{
		console.log('local bot endpoint listening at %s', server.url);
	});
	server.post('/api/messages', bot.connector.listen());
}
else
{
	module.exports = {default: bot.connector.listen()};
}
