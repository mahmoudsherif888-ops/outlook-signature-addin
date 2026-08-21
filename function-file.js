// Corporate Signature - event-based add-in logic
//
// IMPORTANT SCOPE NOTE: Windows New Outlook and Outlook on the web already
// get their signature from a separate, working mechanism (Exchange Online's
// legacy SignatureHtml store, pushed via PowerShell). This add-in exists
// ONLY to cover Outlook for Mac and Outlook Mobile, which do not read that
// same store. To avoid ever double-inserting or conflicting with the
// existing working setup, this code checks Office.context.platform and does
// nothing at all on Windows ("PC") or web ("OfficeOnline") - it only acts
// on "Mac", "iOS", and "Android".
//
// DIRECTORY_URL must point at the signatures-directory.json file generated
// by Export-SignatureDirectory.ps1 and hosted alongside this file.
var DIRECTORY_URL = "https://hr.ain-eg.com/outlook-addin/signatures-directory.json";

function shouldApplyOnThisPlatform() {
  try {
    var platform = Office.context.platform; // "PC","Mac","OfficeOnline","iOS","Android","Universal"
    return platform === "Mac" || platform === "iOS" || platform === "Android";
  } catch (e) {
    return false;
  }
}

function escapeHtml(value) {
  var s = String(value === undefined || value === null ? "" : value);
  return s.replace(/[&<>"]/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
  });
}

function buildSignatureHtml(entry) {
  var parts = [];
  parts.push('<div>');
  if (entry.name)  parts.push('<p style="margin:0"><b>' + escapeHtml(entry.name) + '</b></p>');
  if (entry.title) parts.push('<p style="margin:0">' + escapeHtml(entry.title) + '</p>');
  if (entry.phone) parts.push('<p style="margin:0">' + escapeHtml(entry.phone) + '</p>');
  if (entry.logoUrl) {
    parts.push('<p style="margin:0"><img src="' + entry.logoUrl + '" style="max-width:220px;" alt="logo"></p>');
  }
  parts.push('</div>');
  return parts.join('');
}

function applySignature(event) {
  if (!shouldApplyOnThisPlatform()) {
    event.completed();
    return;
  }

  try {
    var mailbox = Office.context.mailbox;
    var email = ((mailbox.userProfile && mailbox.userProfile.emailAddress) || "").trim().toLowerCase();

    if (!email) {
      event.completed();
      return;
    }

    var controller = (typeof AbortController !== "undefined") ? new AbortController() : null;
    var timeoutId = setTimeout(function () {
      if (controller) { controller.abort(); }
    }, 4000); // never let a slow/unreachable directory server stall compose

    var fetchOptions = { cache: "no-store" };
    if (controller) { fetchOptions.signal = controller.signal; }

    fetch(DIRECTORY_URL, fetchOptions)
      .then(function (resp) {
        clearTimeout(timeoutId);
        if (!resp.ok) { throw new Error("Directory fetch failed: " + resp.status); }
        return resp.json();
      })
      .then(function (directory) {
        var entry = directory[email];
        if (!entry) {
          // No matching entry for this mailbox (e.g. excluded, or not yet
          // in the directory) - do nothing rather than insert a blank/wrong
          // signature.
          event.completed();
          return;
        }
        var html = buildSignatureHtml(entry);
        mailbox.item.body.setSignatureAsync(
          html,
          { coercionType: Office.CoercionType.Html },
          function (asyncResult) {
            // Whether it succeeded or failed, always call completed() so
            // the compose window is never left blocked.
            event.completed();
          }
        );
      })
      .catch(function () {
        event.completed();
      });
  } catch (e) {
    // Never let an unexpected error block the user from composing mail.
    event.completed();
  }
}

// Event-based activation entry points - names must match the
// FunctionName values declared in manifest.xml's LaunchEvents.
function onNewMessageComposeHandler(event) {
  applySignature(event);
}

function onMessageFromChangedHandler(event) {
  applySignature(event);
}

Office.actions = Office.actions || {};
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
