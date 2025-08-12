# Remove Inline Image from Outlook message

This sample is based on the **Manipulate Attachments (Item Compose)** sample included with Script Lab, but with additional code to analyse and remove the inline references from the message body.

To test the code, import [RemoveInlineAttachmentReferences.yml](RemoveInlineAttachmentReferences.yml) into Script Lab.

The new button added with this sample code is **Remove attachment and inline references**.  Other functionality is from the original Script Lab sample.

## Synopsis

As confirmed in our [add-in documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/add-and-remove-attachments-to-an-item-in-a-compose-form#remove-an-attachment), [removeAttachmentAsync](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) will not remove references to attachments from the message body (it will remove the attached file from the message only).

Inline images are represented in a message by an HTML image element with the source (src attribute) pointing to a cid: URL.  This tells the email client that the image itself is an attachment to the message (it will be one of the MIME parts, or in the case of Exchange will have been converted to an attachment in the attachment table of the item), and the cid value contains the reference to the specific image.  This is as per RFCs [2045](https://www.rfc-editor.org/rfc/rfc2045) and [2046](https://www.rfc-editor.org/rfc/rfc2046) (and possibly some others).

Removing inline images from a message is a two stage process:
- remove all references to the image from the HTML source of the message.
- remove the image attachment itself from the message.

If only one of the steps above it completed, then you may observe unexpected results depending on the email client e.g. in Outlook Classic, removing the inline attachment only will not update the message display - so it doesn't look like the attachment is removed.  If you then move off the message and onto it again (to force Outlook to update the message pane), you'll likely see an invalid reference replacing the image that was deleted.

Outlook JS provides a method to remove the attachment from the message, but does not itself implement any logic to remove any references the message body - that is for the add-in developer to implement if needed.

The below function is an example of how to do this.  Note that this is only a simple example with very limited error handling.

```Javascript
function removeInline() {
  // For an inline image, we first need to retrieve the attachment and read its content-id
  var attachmentId = (document.getElementById("attachmentId") as HTMLInputElement).value;
  console.log("Attempting to remove attachment: " + attachmentId);

  Office.context.mailbox.item.getAttachmentsAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Failed to retrieve attachments list.");
      console.error(result.error.message);
      return;
    }
    console.log("Retrieved attachments list.");

    if (result.value.length > 0) {
      for (let i = 0; i < result.value.length; i++) {
        const attachment = result.value[i];
        if (attachment.id == attachmentId) {
          console.log("Found " + attachmentId);
          if (!attachment.isInline) {
            console.error("Attachment is not inline.");
            return;
          }

          // Now we need to remove all references from the message body
          console.log("Retrieving message body.");
          Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
            if (bodyResult.status === Office.AsyncResultStatus.Failed) {
              console.log(`Failed to get message body: ${bodyResult.error.message}`);
              return;
            }
            console.log("Successfully retrieved message body.");
            var messageBody = bodyResult.value;

            // Typical attachment reference:
            // <img width=200 height=200 id="Picture_x0020_1" src="cid:image001.png@01DC0AD0.27C2B7A0">
            //  We need to remove the whole <img> tag, and we identify it by the cid link

            var cidSrc = 'src="cid:' + attachment.name + "@";
            console.log("Searching for references to attachment: " + cidSrc);
            let cidStart = messageBody.indexOf(cidSrc);
            while (cidStart > -1) {
              console.log("Found reference: " + cidSrc);
              let imgStart = messageBody.lastIndexOf("<img", cidStart);
              if (imgStart > -1) {
                console.log("Found <img start.");
                let imgEnd = messageBody.indexOf(">", cidStart);
                if (imgEnd > imgStart) {
                  // Remove <img> element
                  console.log("Determined <img> element: " + messageBody.substring(imgStart, imgEnd + 1));
                  messageBody = messageBody.substring(0, imgStart) + messageBody.substring(imgEnd + 1);
                }
              }
              cidStart = messageBody.indexOf(cidSrc, cidStart + 1);
            }

            console.log("Updated body:");
            console.log(messageBody);

            Office.context.mailbox.item.body.setAsync(
              messageBody,
              { coercionType: Office.CoercionType.Html },
              (result) => {
                if (result.status == Office.AsyncResultStatus.Failed) {
                  console.log("Failed to update message body.");
                  return;
                }

                // Now remove attachment
                console.log("Removing inline attachment: " + attachmentId);
                console.log(attachment);
                Office.context.mailbox.item.removeAttachmentAsync(attachmentId, (result) => {
                  if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error(result.error.message);
                    return;
                  }
                  console.log(`Attachment removed successfully.`);
                });
              },
            );
          });
        }
      }
    } else {
      console.log("No attachments on this message.");
    }
    console.log("Finished removing inline attachment.");
  });
}
```