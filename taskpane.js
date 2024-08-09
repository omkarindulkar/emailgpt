/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     document.getElementById("run").onclick = run;
//   }
// });

// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */

//   const item = Office.context.mailbox.item;
//   let insertAt = document.getElementById("item-subject");
//   let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
//   insertAt.appendChild(label);
//   insertAt.appendChild(document.createElement("br"));
//   insertAt.appendChild(document.createTextNode(item.subject));
//   insertAt.appendChild(document.createElement("br"));
// }

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  const item = Office.context.mailbox.item;

  // Get the email body
  item.body.getAsync("html", async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;

      try {
        // Send the email body to the backend server
        const response = await fetch('https://leadscoringv2.azurewebsites.net/suggest-reply', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ emailThread: emailBody })
        });

        // Process the response from the backend
        const data = await response.json();
        console.log(data);
        const suggestedReply = data.suggestedReply;

        // Display the suggested reply in the task pane
        document.getElementById("suggested-reply").innerText = suggestedReply;
      } catch (error) {
        console.error('Error fetching suggested reply:', error);
        document.getElementById("suggested-reply").innerText = "Error fetching suggested reply.";
      }
    } else {
      console.error('Failed to get email body:', result.error);
      document.getElementById("suggested-reply").innerText = "Failed to get email body.";
    }
  });
}
