/* ============================================================
   Gekonny Subject Builder — Smart Alert (OnMessageSend)
   Accepts BOTH formats:
     new:  [RFI] 22 East 10 - question [GC_2026_0007]
     old:  [GC_2026_0007] [RFI] question
     bid:  [RFP GC] project name
   ============================================================ */

Office.onReady(function () {});

var RE_CODE   = /\[(SUB|GC)_\d{4}_\d{4}\]/i;      // project code anywhere
var RE_LEAD   = /^\s*\[[^\]]+\]/;                  // starts with a [tag]
var RE_NEWBID = /\[\s*RFP\s+(GC|SUB)\s*\]/i;

function subjectLooksTagged(subject) {
  if (!subject) { return false; }
  var cleaned = subject.replace(/^(\s*(re|fw|fwd)\s*:\s*)+/i, "");
  if (RE_NEWBID.test(cleaned)) { return true; }
  return RE_CODE.test(cleaned) && RE_LEAD.test(cleaned);
}

/* People who should never see the send-time prompt.

   The add-in itself stays for them — panel, ribbon button, everything.
   Only the interruption is dropped, because for someone who does not file
   work through Monday the prompt is pure noise on every unrelated email.
   Removing the whole add-in from their mailbox would work too, but it
   takes away a tool to silence a nag.

   Addresses, lower case, one per line. */
var ALERT_EXEMPT = [
];

function alertIsExemptForMe() {
  try {
    var me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
    if (!me) { return false; }
    for (var i = 0; i < ALERT_EXEMPT.length; i++) {
      if (String(ALERT_EXEMPT[i]).toLowerCase() === me) { return true; }
    }
  } catch (e) { /* no profile — treat as not exempt */ }
  return false;
}

function onMessageSendHandler(event) {
  if (alertIsExemptForMe()) {
    event.completed({ allowEvent: true });
    return;
  }

  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }

    var subject = result.value || "";

    if (subjectLooksTagged(subject)) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage:
          "This subject is not formatted for Monday automation, so no task will be created.\n\n" +
          "Use the Subject Builder (Gekonny tab) to format it as:\n" +
          "  [RFI] 22 East 10 - description [GC_2026_0007]  — for an existing project\n" +
          "  [RFP GC] project name  — for a new bid\n\n" +
          "Send anyway if this email is not project-related.",
        cancelLabel: "Fix subject",
        commandId: "openSubjectBuilderButton"
      });
    }
  });
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
