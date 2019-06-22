const xapi = require('xapi');

// Global constants & variables
const supportURI = 'support@company.com'
var systemInfo = {
  productId: '',
  systemName: '',
  softwareVersion: '',
};


// Functions
function post(msg) {
  // Replace with your bot token
  const token = "Zjg1ZjQ2ZGUtMDQyZC00Njk1LTg4NmQtMWQxMzE2ZjYzZTA3MGJmMjBhZjMtMWU5_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f"
  // replace with a space your bot is part of
  const roomId = 'Y2lzY29zcGFyazovL3VzL1JPT00vYjNiYWY3ODAtMWIyNy0xMWU5LThhMzEtN2ZkOGM4YjBhNzI1'
  // Post message
  let payload = {"markdown": msg, "roomId": roomId}
  xapi.command( 'HttpClient Post',
    {
      Header: ["Content-Type: application/json", "Authorization: Bearer " + token],
      Url: "https://api.ciscospark.com/v1/messages",
      AllowInsecureHTTPS: "True"
    },
    JSON.stringify(payload))
    .then((response) => {
      if (response.StatusCode == 200) {
        console.log("message pushed to Webex Teams")
        return
      }
    })
    .catch((err) => {
      console.log("failed with err: " + err.message)
    })
}

function init() {
  console.log('Report issues macro initializing...')
  // This needs to be set to allow HTTP Post
  xapi.config.set('HttpClient Mode', 'On');
  xapi.config.set('HttpClient AllowInsecureHTTPS', 'True');
  // Retrieves the codec info
  xapi.status.get('SystemUnit ProductId').then((value) => {
    systemInfo.productId = value;
  });
  xapi.status.get('SystemUnit Software Version').then((value) => {
    systemInfo.softwareVersion = value;
  });
  xapi.status.get('UserInterface ContactInfo Name').then((value) => {
    systemInfo.systemName = value;
  });  
  console.log('Report issues macro listening...');
}


// Event listeners
xapi.event.on('UserInterface Extensions Widget Action', (event) => {

    // Call support
    if (event.WidgetId == 'btnCallSupport' && event.Type == 'clicked') {
        xapi.command('Dial', {Number: supportURI});
        console.log('Call Support in progress...');
    }

    // Post to Teams
    if (event.WidgetId == 'btnPostMessage' && event.Type == 'clicked') {
        xapi.command("UserInterface Message Prompt Display", {
          Title: 'Report room issue',
          Text: 'Please select what the problem area is:',
          FeedbackId: 'roomfeedback',
          'Option.1':'Cleanliness',
          'Option.2':'Technical issues with Audio/Video',
          'Option.3': 'Other'
        })
    }
});

xapi.event.on('UserInterface Message Prompt Response', (event) => {
    if (event.FeedbackId == 'roomfeedback') {
      console.log('Selected option: ' + event.OptionId)
      var message;
      if (event.OptionId == '1') {
        post('There is an issue in ' + systemInfo.systemName + ', the room is dirty!');
      }
      if (event.OptionId == '2') {
        message = 'There is an **audio/video** issue in ' + systemInfo.systemName;
        message += '<br/> *Device details:*'; 
        message += '<br/> - Type: **' + systemInfo.productId + '**';
        message += '<br/> - Version: **' + systemInfo.softwareVersion + '**';
        post(message);
      }
      if (event.OptionId == '3') {
        message = 'There is an issue in ' + systemInfo.systemName;
        message += '<br/> *Device details:*'; 
        message += '<br/> - Type: **' + systemInfo.productId + '**';
        post(message);
      }
      xapi.command('UserInterface Extensions Panel Close');
      xapi.command("UserInterface Message Alert Display", {
          Title: 'Feedback receipt',
          Text: 'Thank you for you feedback! Have a great day!',
          Duration: 5
    });
  }
});


init();
