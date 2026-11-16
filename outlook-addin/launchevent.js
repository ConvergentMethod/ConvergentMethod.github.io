/**
 * Convergent Method - Microsoft 365 Outlook Add-in
 * Automatically applies the official corporate HTML signature and disclaimer.
 */

/* global Office */

// Configuration & Team Directory
var TEAM_DIRECTORY = {
  "markst@convergentmethod.com": {
    name: "Mark Stemmans, P.E.",
    title: "Construction Program Manager",
    email: "MarkSt@convergentmethod.com",
    phone: "(281) 979-1336",
    phoneLink: "+12819791336"
  },
  "amandast@convergentmethod.com": {
    name: "Amanda Stemmans, MS, APRN, FNP-C",
    title: "Chief Operations Officer",
    email: "AmandaSt@convergentmethod.com",
    phone: "(281) 979-2162",
    phoneLink: "+12819792162"
  }
};

// If true, forces Mark Stemmans' signature across all accounts regardless of logged-in user
var FORCE_MARK_SIGNATURE = false;

Office.onReady(function () {});

function onNewMessageComposeHandler(event) {
  applySignature(event);
}

function onNewAppointmentOrganizerHandler(event) {
  applySignature(event);
}

function buildSignatureHtml(user) {
  return (
    '<!-- Microsoft 365 Auto Signature - Convergent Method, LLC -->' +
    '<div style="font-family: \'Aptos\', Calibri, Arial, sans-serif; font-size: 13px; color: #333333; line-height: 1.4; padding-top: 16px;">' +
      '<table width="560" cellpadding="0" cellspacing="0" border="0" style="width: 560px; min-width: 560px; max-width: 560px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-family: \'Aptos\', Calibri, Arial, sans-serif;">' +
        '<tr>' +
          '<td valign="middle" width="145" style="width: 145px; min-width: 145px; max-width: 145px; padding: 0 16px 0 0; border-right: 2px solid #6FA0DB; vertical-align: middle;">' +
            '<a href="https://convergentmethod.com" target="_blank" style="text-decoration: none;">' +
              '<img src="https://convergentmethod.com/media/CMFullColor.png" alt="Convergent Method Logo" width="145" border="0" style="display: block; width: 145px; max-width: 145px; height: auto; border: 0; outline: none; text-decoration: none;" />' +
            '</a>' +
          '</td>' +
          '<td valign="middle" width="395" style="width: 395px; min-width: 395px; padding: 0 0 0 16px; vertical-align: middle; white-space: nowrap;">' +
            '<div style="font-family: \'Century Gothic\', CenturyGothic, \'Apple Gothic\', \'Avenir Next\', Avenir, \'Segoe UI\', Arial, sans-serif; font-size: 17px; font-weight: 700; color: #6FA0DB; letter-spacing: 0.2px; line-height: 1.2; margin: 0 0 3px 0; white-space: nowrap;">' +
              escapeHtml(user.name) +
            '</div>' +
            '<div style="font-family: \'Century Gothic\', CenturyGothic, \'Apple Gothic\', \'Avenir Next\', Avenir, \'Segoe UI\', Arial, sans-serif; font-size: 13px; font-weight: 600; color: #222222; line-height: 1.3; margin: 0 0 2px 0; white-space: nowrap;">' +
              escapeHtml(user.title) +
            '</div>' +
            '<div style="font-family: \'Aptos\', Calibri, Arial, sans-serif; font-size: 13px; font-weight: 400; color: #555555; line-height: 1.3; margin: 0 0 8px 0; white-space: nowrap;">' +
              'Convergent Method, LLC' +
            '</div>' +
            '<div style="font-family: \'Aptos\', Calibri, Arial, sans-serif; font-size: 12px; line-height: 1.4; margin: 0 0 2px 0; white-space: nowrap;">' +
              '<span style="font-weight: 600; color: #333333;">E:&nbsp;</span>' +
              '<a href="mailto:' + encodeURIComponent(user.email) + '" style="color: #6FA0DB; text-decoration: none; font-weight: 500;">' +
                '<span style="color: #6FA0DB; text-decoration: none; font-weight: 500;">' + escapeHtml(user.email) + '</span>' +
              '</a>' +
            '</div>' +
            '<div style="font-family: \'Aptos\', Calibri, Arial, sans-serif; font-size: 12px; line-height: 1.4; margin: 0; white-space: nowrap;">' +
              '<span style="font-weight: 600; color: #333333;">P:&nbsp;</span>' +
              '<a href="tel:' + user.phoneLink + '" style="color: #444444; text-decoration: none;">' +
                '<span style="color: #444444; text-decoration: none;">' + escapeHtml(user.phone) + '</span>' +
              '</a>' +
            '</div>' +
          '</td>' +
        '</tr>' +
        '<tr><td colspan="2" height="16" style="height: 16px; font-size: 0; line-height: 0; padding: 0;">&nbsp;</td></tr>' +
        '<tr>' +
          '<td colspan="2" style="padding: 0; margin: 0; font-family: \'Aptos\', Calibri, Arial, sans-serif;">' +
            '<div style="font-size: 10px; line-height: 14px; color: #6FA0DB; font-style: italic; padding-bottom: 6px;">' +
              'Please do not print this e-mail unless necessary' +
            '</div>' +
            '<div style="font-size: 10px; line-height: 14px; color: #888888;">' +
              '<strong style="color: #666666;">Notice:</strong> This message is confidential, is intended only for the named recipient(s) and may contain information that is privileged or exempt from disclosure under applicable law. If you are not the intended recipient(s), you are notified that the dissemination, distribution or copying of this message is strictly prohibited. If you received this message and are not an intended recipient, please delete it from your computer.' +
            '</div>' +
          '</td>' +
        '</tr>' +
      '</table>' +
    '</div>'
  );
}

function resolveCurrentUserInfo() {
  if (FORCE_MARK_SIGNATURE) {
    return TEAM_DIRECTORY["markst@convergentmethod.com"] || Object.values(TEAM_DIRECTORY)[0];
  }
  try {
    var profile = Office.context.mailbox.userProfile;
    var email = (profile && profile.emailAddress) ? profile.emailAddress.trim().toLowerCase() : "";
    var displayName = (profile && profile.displayName) ? profile.displayName.trim() : "";

    if (email && TEAM_DIRECTORY[email]) return TEAM_DIRECTORY[email];
    if (email.indexOf("mark") !== -1 || displayName.toLowerCase().indexOf("mark") !== -1) {
      return TEAM_DIRECTORY["markst@convergentmethod.com"];
    }
    if (email.indexOf("amanda") !== -1 || displayName.toLowerCase().indexOf("amanda") !== -1) {
      return TEAM_DIRECTORY["amandast@convergentmethod.com"];
    }
    return {
      name: displayName || "Convergent Method Team",
      title: "Convergent Method, LLC",
      email: email || "info@convergentmethod.com",
      phone: "(281) 979-1336",
      phoneLink: "+12819791336"
    };
  } catch (e) {
    return TEAM_DIRECTORY["markst@convergentmethod.com"] || Object.values(TEAM_DIRECTORY)[0];
  }
}

function applySignature(event) {
  try {
    var item = Office.context.mailbox.item;
    if (item && item.body && typeof item.body.setSignatureAsync === "function") {
      var user = resolveCurrentUserInfo();
      var signatureHtml = buildSignatureHtml(user);
      item.body.setSignatureAsync(
        signatureHtml,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
          event.completed({ allowEvent: true });
        }
      );
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (err) {
    event.completed({ allowEvent: true });
  }
}

function insertSignatureButtonHandler(event) {
  applySignature(event);
}

Office.actions.associate("insertSignatureButtonHandler", insertSignatureButtonHandler);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onNewAppointmentOrganizerHandler", onNewAppointmentOrganizerHandler);
