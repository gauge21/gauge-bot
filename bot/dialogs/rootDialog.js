const { ComponentDialog } = require("botbuilder-dialogs");
const { TurnContext, ActionTypes, CardFactory, TextFormatTypes } = require("botbuilder");

class RootDialog extends ComponentDialog {
  constructor(id) {
    super(id);
  }

  async onBeginDialog(innerDc, options) {
    const result = await this.triggerCommand(innerDc);
    if (result) {
      return result;
    }

    return await super.onBeginDialog(innerDc, options);
  }

  async onContinueDialog(innerDc) {
    return await super.onContinueDialog(innerDc);
  }

  async triggerCommand(innerDc) {
    const removedMentionText = TurnContext.removeRecipientMention(
      innerDc.context.activity,
      innerDc.context.activity.recipient.id
    );
    let text = "";
    if (removedMentionText) {
      text = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim(); // Remove the line break
    }

    if (innerDc.context.activity.textFormat !== TextFormatTypes.Plain) {
      return await innerDc.cancelAllDialogs();
    }

    switch (text) {
      case "show": {
        if (innerDc.context.activity.conversation.isGroup) {
          await innerDc.context.sendActivity(
            `Sorry, currently TeamsFX SDK doesn't support Group/Team/Meeting Bot SSO. To try this command please install this app as Personal Bot and send "show".`
          );
          return await innerDc.cancelAllDialogs();
        }
        break;
      }
      case "intro": {
        const cardButtons = [
          {
            type: ActionTypes.ImBack,
            title: "Show profile",
            value: "show",
          },
        ];
        const card = CardFactory.heroCard("Introduction", null, cardButtons, {
          text: `How are you feeling?`,
        });
        await innerDc.context.sendActivity({ attachments: [card] });
        return await innerDc.cancelAllDialogs();
      }
      default: {
        const cardButtons = [
          {
            type: ActionTypes.ImBack,
            title: "Show introduction card",
            value: "intro",
          },
        ];
        const card = CardFactory.heroCard("", null, cardButtons, {
          text: `This is a hello world Bot built with Microsoft Teams Framework, which is designed for illustration purposes. This Bot by default will not handle any specific question or task.<br>Please type <strong>intro</strong> to see the introduction card.`,
        });
        await innerDc.context.sendActivity({ attachments: [card] });
        return await innerDc.cancelAllDialogs();
      }
    }
  }
}

module.exports.RootDialog = RootDialog;
