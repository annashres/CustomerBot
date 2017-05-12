# MAAV-1 Bot

This bot has been created using [Microsoft Bot Framework](https://dev.botframework.com), and scaffolded using the [Bot Builder Yeoman generator](https://github.com/GeekTrainer/generator-botbuilder).

## Purpose
MAAV-1 is designed to record notes and insights from the conversations Microsoft employees have with customers. The goal is to make it easier to disseminate knowledge about Microsoft customers and win as a team.

## Bot Deployment
**Deploy bot automatically on Azure**  
[![Deploy to Azure](http://azuredeploy.net/deploybutton.svg)](https://azuredeploy.net/)

**Deploy bot on local machine**  
1. Make sure you have [node.js and npm](https://nodejs.org/en/) installed on your local machine. 

2. Create a new file in the project directory called `.env`. This file will be used to setup environment variables that the bot will use to run.

3. Copy the contents of the `.env.defaults` file into the `.env` file.

4. Modify the `.env` file to add the following: `NODE_ENV = development`. This sets up your local machine as a dev environment.

5. Run `npm install` in the bot directory to install required dependencies.

6. Open a command prompt window in the project directory. You can now start the bot locally by running `node app.js`

### Testing/Using the bot

**On Azure**  
Use the chat window built into the Azure bot service to send and receive messages from MAAV-1

**On local machine**  
* Download the [bot framework emulator](https://download.botframework.com/bf-v3/tools/emulator/publish.htm)
* Launch the emulator and use the chat window to send and receive messages from MAAV-1

### Structure

`app.js` references the bot and starts a [Restify](http://restify.com/) server. `bot.js` has a simple multi-turn dialog which sends the name and description of the bot, and then asks the user for their name.

`host.json` and `function.json` are required configuration files needed to setup the bot framework on an Azure function app.

`package.json` holds metadata relevant to the project including its list of dependencies.

`azuredeploy.json` and `azuredeploy.parameters.json` hold the Azure Resource Manager template and configuration info needed to deploy the MAAV-1 bot to Azure.


### Configuring the bot

The template uses [dotenv](https://github.com/motdotla/dotenv) for managing application settings.