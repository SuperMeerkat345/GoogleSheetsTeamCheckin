function PingScouter(name, team, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const secrets = ss.getSheetByName("SENSITIVE")

  const scriptProps = PropertiesService.getScriptProperties();
  const SLACKTOKEN = scriptProps.getProperty('SLACKTOKEN');
  const CHANNELID = scriptProps.getProperty('CHANNELID');

  const data = secrets.getRange("B2:C" + secrets.getLastRow()).getValues();
  
  let slackUserId = "";
  
  // Look for the name in the array
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === name) {
      slackUserId = data[i][1];
      break;
    }
  }

  // Handle case where name isn't found
  if (!slackUserId) {
    console.error("User not found in DB: " + name);
    return;
  }
  const url = "https://slack.com/api/chat.postMessage";
  
  const payload = {
    "channel": CHANNELID,
    "text": "New Task Assigned!", // Fallback text for notifications
    "username": "TEAM CHECKIN BOT",
    "icon_url": "https://cdn-icons-png.flaticon.com/512/1632/1632670.png",
    "blocks": [
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": "🚨 New Task Assigned 🚨",
          "emoji": true
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "<@" + slackUserId + ">, you've got a new assignment."
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "section",
        "fields": [
          {
            "type": "mrkdwn",
            "text": "*Team:*\n" + team
          },
          {
            "type": "mrkdwn",
            "text": "*Status:*\nPending"
          }
        ]
      }, 
      {
        "type": "section",
        "fields": [
          {
            "type": "mrkdwn",
            "text": "*Description:*\n" + description
          },
          {
            "type": "mrkdwn",
            "text": `*Pit Location*\nhttps://frc.nexus/en/event/2026ohcl/team/${team}/map`
          }
        ]
      },
      {
        "type": "context",
        "elements": [
          {
            "type": "mrkdwn",
            "text": "Sent from Google Sheets Scouter Bot"
          }
        ]
      }
    ]
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + SLACKTOKEN
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const resContent = JSON.parse(response.getContentText());
    
    if (resContent.ok) {
      console.log("Message sent successfully!");
    } else {
      console.error("Slack Error: " + resContent.error);
    }
  } catch (e) {
    console.error("Connection Error: " + e.toString());
  }
}

function test2() {
  PingScouter("Artiom", "695", "Test")
}
