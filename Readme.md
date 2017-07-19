# MAAV-1 Bot

This bot has been created using [Microsoft Bot Framework](https://dev.botframework.com), and scaffolded using the [Bot Builder Yeoman generator](https://github.com/GeekTrainer/generator-botbuilder).

## Purpose
MAAV-1 is designed to record notes and insights from the conversations Microsoft employees have with customers. The goal is to make it easier to disseminate knowledge about Microsoft customers and win as a team.

## High-level Architecture
![MAAV-1 Architecture](Maav1Architecture.JPG)

## Deploy a new Bot instance
1. Navigate to the [Azure portal](https://portal.azure.com). Create a new 'Bot service' instance.

2. Follow the instructions on the Bot blade to create a Microsoft APP ID and Password for the bot service instance.

3. Choose the NodeJS template for the bot.

4. Select the 'Basic' template for initial bot code.

5. Navigate back to the Azure portal and create a new SQL database. This database will be used to store login pin information.

The bot service instance has now been deployed with initial default code for basic interactive communication.

## Upload MAAV-1 code to bot 
1. Navigate to the 'Settings' tab of the bot instance.

2. Scroll to the 'Continuous Integration' section and click **Configure Continuous Integration**.

3. Set up the bot to pull code from the 'feature-AzureBot' branch of the [CustomerBot Github repository](https://github.com/annashres/CustomerBot).

4. Navigate back to the 'Settings' tab of the bot instance and scroll down to the 'Application Settings' section. Click the **Open** button.

5. Scroll down to the 'Application settings' section of the blade and add the following environment variables (fill in applicable values):
	* MicrosoftAppId
	* MicrosoftAppPassword
	* NODE_ENV
	* DB_ADMIN
	* DB_PASSWORD
	* DB_SERVER 
	* DB_NAME
	* FEEDBACKDB_NAME 
	* FEEDBACKDB_ADMIN
	* FEEDBACKDB_PASSWORD
	* FEEDBACKDB_SERVER 
	* DashboardUrl 
	
	Save the environment variables once you have finished editing.

6. Optional: Clone the [CustomerBot Github repository](https://github.com/annashres/CustomerBot/tree/feature-AzureBot) to your local machine. This will allow you to test changes locally.


### Testing/Using the bot

**On Azure**  
Use the chat window built into the Azure bot service to send and receive messages from MAAV-1

**On local machine**  

* Make sure you have downloaded the source code for the MAAV-1 bot from the [CustomerBot Github repository](https://github.com/annashres/CustomerBot/tree/feature-AzureBot).
* Download the [bot framework emulator](https://download.botframework.com/bf-v3/tools/emulator/publish.htm)
* MAAV-1 uses [dotenv](https://github.com/motdotla/dotenv) for managing application settings. You will need to create and edit a '.env' file to manage environment variables needed by the bot.
	* Copy the contents of the '.env.defaults' file into a new '.env' file.
	* Edit the 'NODE_ENV' variable to have a value 'development'. This sets up the bot to run on your local machine. 
	* Fill out the remaining required environment variables. You can omit filling out values for the 'MicrosoftAppId' and 'MicrosoftAppPassword' variables.
* Open a command prompt window in the project directory. Navigate to the 'messages' folder. You can start the bot locally by running `node index.js`.
* Launch the emulator and use the chat window to send and receive messages from MAAV-1

### Structure

`PostDeployScripts` contains a bash script that setups the bot environment with appropriate environment variables and node modules each time it is deployed on Azure.

`host.json` and `messages/function.json` are required configuration files needed to setup the bot framework on an Azure function app.

`messages/.vscode` is default configuration file needed to the debug the bot using VSCode.

`messages/.env.defaults` is the template for the `.env` file that should be created if testing/working on a local machine.

`messages/bot.js` is the core bot code. It uses the Microsoft bot framework and the bot builder library to collect information.

`messages/botdisplay.js` is the channel display module for the bot. It abstracts display details for displaying conversation details in different channels.

`messages/index.js` exposes the bot via an Azure function app. Alternatively, it starts a [Restify](http://restify.com/) server on a local machine. 

`messages/package.json` holds metadata relevant to the project including its list of dependencies.





