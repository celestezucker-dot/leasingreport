/* function-file.js
   Outlook Add-in OnMessageSend handler.
   Blocks sending if the subject contains "Daily Leasing Report"
   and there is no .pdf attachment.
*/

Office.onReady(() => {
  // Office.onReady must be called, even though this runtime has no UI.
});

function checkLeasingReportPdf(event) {
  const item = Office.context.mailbox.item;

  item.subject.getAsync((subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      // If we can't even read the subject, don't block the send.
      event.completed({ allowEvent: true });
      return;
    }

    const subject = subjectResult.value || "";
    const requiresPdf = subject.toLowerCase().indexOf("daily leasing report") > -1;

    if (!requiresPdf) {
      event.completed({ allowEvent: true });
      return;
    }

    item.getAttachmentsAsync((attachmentsResult) => {
      if (attachmentsResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: true });
        return;
      }

      const attachments = attachmentsResult.value || [];
      const hasPdf = attachments.some((att) => {
        const name = (att.name || "").toLowerCase();
        return name.endsWith(".pdf");
      });

      if (hasPdf) {
        event.completed({ allowEvent: true });
      } else {
        event.completed({
          allowEvent: false,
          errorMessage:
            "This email's subject contains 'Daily Leasing Report' but no PDF is attached. Please attach the PDF before sending."
        });
      }
    });
  });
}

// Register the function so the runtime can find it.
Office.actions.associate("checkLeasingReportPdf", checkLeasingReportPdf);
