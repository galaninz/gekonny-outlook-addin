/* ============================================================
   Gekonny Subject Builder — Smart Alert (OnMessageSend)
   Checks the subject when the user hits Send.
   Soft warning only: the user can still choose to send.
   ============================================================ */

Office.onReady(function () {
  // handler registration happens via the associate() call below
});

/* ---- Subject validation ----
   A subject is considered "tagged" if it matches EITHER:
   (1) existing project:  [SUB_2026_0001] [RFI] ...
   (2) new bid intake:    RFP GC - ...   /   RFP SUB - ...
---------------------------------------------------------------- */
var RE_EXISTING = /^\s*\[(SUB|GC)_\d{4}_\d{4}\]\s*\[[^\]]+\]/i;
var RE_NEWBID   = /^\s*RFP\s+(GC|SUB)\s*-/i;

function subjectLooksTagged(subject) {
  if (!subject) { return false; }
  // strip leading RE: / FW: / FWD: prefixes before checking
  var cleaned = subject.replace(/^(\s*(re|fw|fwd)\s*:\s*)+/i, "");
  return RE_EXISTING.test(cleaned) || RE_NEWBID.test(cleaned);
}

/* ---- The OnMessageSend handler ---- */
function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      // if we cannot read the subject, do not block the user
      event.completed({ allowEvent: true });
      return;
    }

    var subject = result.value || "";

    if (subjectLooksTagged(subject)) {
      // subject is properly formatted — let the email go
      event.completed({ allowEvent: true });
    } else {
      // soft warning — user may still send
      event.completed({
        allowEvent: false,
        errorMessage:
          "This subject is not formatted for Monday automation, so no task will be created.\n\n" +
          "Use the Subject Builder (Gekonny tab) to format it as:\n" +
          "  [SUB_2026_0001] [RFI] description     — for an existing project\n" +
          "  RFP GC - project name                 — for a new bid\n\n" +
          "Send anyway if this email is not project-related.",
        cancelLabel: "Fix subject",
        commandId: "openSubjectBuilderButton"
      });
    }
  });
}

/* ---- Register the handler with Office ---- */
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
