const { TeamsActivityHandler, CardFactory, TurnContext} = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const ACData = require("adaptivecards-templating");
const {loadConfiguration, DefaultTediousConnectionConfiguration} = require("@microsoft/teamsfx");
const {Connection, Request} = require("tedious");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // record the likeCount
    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = this.renderAdaptiveCard(rawWelcomeCard);
          await context.sendActivity({ attachments: [card] });
          break;
        }
        case "learn": {
          // this.likeCountObj.likeCount = 0;
          // const card = this.renderAdaptiveCard(rawLearnCard, this.likeCountObj);
          // await context.sendActivity({ attachments: [card] });
          loadConfiguration();
          const connection = await getSQLConnection();
          const query = "select system_user as u, sysdatetime() as t";
          const result = await execQuery(query, connection);
          await context.sendActivities(result[0][0]);
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    async function getSQLConnection() {
      const sqlConnectConfig = new DefaultTediousConnectionConfiguration();
      const config = await sqlConnectConfig.getConfig();
      const connection = new Connection(config);
      return new Promise((resolve, reject) => {
        connection.on("connect", (error) => {
          if (error) {
            console.log(error);
            reject(connection);
          }
          resolve(connection);
        });
      });
    }
    
    async function execQuery(query, connection) {
      return new Promise((resolve, reject) => {
        const res = [];
        const request = new Request(query, (err) => {
          if (err) {
            throw err;
          }
        });
    
        request.on("row", (columns) => {
          const row = [];
          columns.forEach((column) => {
            row.push(column.value);
          });
          res.push(row);
        });
        request.on("requestCompleted", () => {
          resolve(res);
        });
        request.on("error", () => {
          console.error("SQL execQuery failed");
          reject(res);
        });
        connection.execSql(request);
      });
    }

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = this.renderAdaptiveCard(rawWelcomeCard);
          await context.sendActivity({ attachments: [card] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(context, invokeValue) {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = this.renderAdaptiveCard(rawLearnCard, this.likeCountObj);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [card],
      });
      return { statusCode: 200 };
    }
  }

  // Bind AdaptiveCard with data
  renderAdaptiveCard(rawCardTemplate, dataObj) {
    const cardTemplate = new ACData.Template(rawCardTemplate);
    const cardWithData = cardTemplate.expand({ $root: dataObj });
    const card = CardFactory.adaptiveCard(cardWithData);
    return card;
  }

}


module.exports.TeamsBot = TeamsBot;
