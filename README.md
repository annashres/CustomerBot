# MAAV-1 Bot

This bot has been created using [Microsoft Bot Framework](https://dev.botframework.com), and scaffolded using the [Bot Builder Yeoman generator](https://github.com/GeekTrainer/generator-botbuilder).

This bot is designed to record notes and insights from the conversations Microsoft employees have with customers. The goal is to make it easier to disseminate knowledge about Microsoft customers and win as a team.

[![Deploy to Azure](http://azuredeploy.net/deploybutton.png)](https://azuredeploy.net/)

## Getting Started on Azure
Deploy MAAV-1, by clicking the 'Deploy to Azure' button above.

## Getting Started on Local Machine
When you clone locally, make to sure to create a '.env' in the root directory. This file will store the environment variables needed to run the bot.
NODE_ENV = development

### Structure

`app.js` references the bot and starts a [Restify](http://restify.com/) server. `bot.js` has a simple multi-turn dialog which sends the name and description of the bot, and then asks the user for their name.

`host.json` and `function.json` are required configuration files needed to setup the bot framework on an Azure function app.

`package.json` holds metadata relevant to the project including its list of dependencies.

`azuredeploy.json` and `azuredeploy.parameters.json` hold the Azure Resource Manager template and configuration info needed to deploy the MAAV-1 bot to Azure.

### Running the bot

**Locally**

```
node app.js
```

**On Azure**
Deploy the app and use the provided chat emulator to test out the bot


### Configuring the bot

The template uses [dotenv](https://github.com/motdotla/dotenv) for managing application settings.