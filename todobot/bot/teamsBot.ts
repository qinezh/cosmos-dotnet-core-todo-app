import { TeamsActivityHandler, CardFactory, TurnContext, Attachment, AdaptiveCardInvokeValue, AdaptiveCardInvokeResponse} from "botbuilder";
import axios from "axios";
import * as https from "https";

const ACData = require("adaptivecards-templating");
const endpoint = "https://localhost:5001";

enum Command {
  add = "add todo",
  list = "list todo"
}

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

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

      const instance = axios.create({
        httpsAgent: new https.Agent({  
          rejectUnauthorized: false
        })
      });

      // Trigger command by IM text
      if (txt.startsWith(Command.add)) {
        const todo = txt.slice(txt.indexOf(Command.add)+Command.add.length).trim();
        // create todo
        const params = new URLSearchParams();
        params.append('Name', todo);
        params.append('Completed', "false");

        await instance.post(`${endpoint}/item/CreateAPI`, params);

        await context.sendActivity(`new todo ${todo} created.`);
      } else if (txt == Command.list) {
        // list todo
        const response = await instance.get(`${endpoint}/item/ListAPI`);
        const todos = []
        for (const item of response.data) {
          todos.push(item.name);
        }
        await context.sendActivity(todos.join('\n'));
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity("welcome");
          break;
        }
      }
      await next();
    });
  }

  // Bind AdaptiveCard with data
  renderAdaptiveCard(rawCardTemplate: any, dataObj?: any): Attachment {
    const cardTemplate = new ACData.Template(rawCardTemplate);
    const cardWithData = cardTemplate.expand({ $root: dataObj });
    const card = CardFactory.adaptiveCard(cardWithData);
    return card;
  }

}

