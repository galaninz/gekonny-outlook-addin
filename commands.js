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

function onMessageSendHandler(event) {
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
