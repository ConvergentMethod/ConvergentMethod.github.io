/**
 * Convergent Method - Microsoft 365 Outlook Add-in
 * Automatically applies the official corporate HTML signature and disclaimer.
 *
 * SECURITY & PRIVACY POLICY:
 * - This add-in NEVER reads email body contents, subject lines, recipients, or attachments.
 * - It ONLY accesses the sender's own identity (userProfile.displayName and userProfile.emailAddress)
 *   to determine the appropriate corporate signature.
 * - Air-gapped: Zero outbound network calls (no fetch, no XHR, no WebSocket, no third-party tracking).
 */

/* global Office */

// Configuration & Team Directory (Immutable)
var TEAM_DIRECTORY = Object.freeze({
  "ethanstemmans@gmail.com": Object.freeze({
    name: "Ethan Stemmans",
    title: "Electrical Engineering Assistant",
    company: "Convergent Method, LLC",
    email: "ethanstemmans@gmail.com",
    phone: "(281) 678-6643",
    phoneLink: "2816786643",
    hasLogo: true,
    hasDisclaimer: true,
    hasEco: true,
    hasExtraDetail: false,
    extraLabel: "Dealer #:",
    extraValue: "",
    signatures: Object.freeze([
      Object.freeze({
        id: "sig_1",
        name: "Personal Email",
        fullName: "Ethan Stemmans",
        title: "Electrical Engineering Assistant",
        company: "Convergent Method, LLC",
        email: "ethanstemmans@gmail.com",
        phone: "(281) 678-6643",
        phoneLink: "2816786643",
        hasLogo: true,
        logo: "https://convergentmethod.com/media/CMFullColor.png",
        color: "#6fa0db",
        hasDisclaimer: true,
        hasEco: true,
        hasExtraDetail: false,
        extraLabel: "Dealer #:",
        extraValue: ""
      })
    ])
  }),
  "markst@convergentmethod.com": Object.freeze({
    name: "Mark Stemmans, P.E.",
    title: "Construction Program Manager",
    company: "Convergent Method, LLC",
    email: "MarkSt@convergentmethod.com",
    phone: "(281) 979-1336",
    phoneLink: "2819791336",
    hasLogo: true,
    hasDisclaimer: true,
    hasEco: true,
    hasExtraDetail: false,
    extraLabel: "Dealer #:",
    extraValue: "",
    signatures: Object.freeze([
      Object.freeze({
        id: "sig_1794890126652_1",
        name: "Project",
        fullName: "Mark Stemmans, P.E.",
        title: "Construction Program Manager",
        company: "Convergent Method, LLC",
        email: "MarkSt@convergentmethod.com",
        phone: "(281) 979-1336",
        phoneLink: "2819791336",
        hasLogo: true,
        logo: "https://convergentmethod.com/media/CMFullColor.png",
        color: "#6fa0db",
        hasDisclaimer: true,
        hasEco: true,
        hasExtraDetail: false,
        extraLabel: "Dealer #:",
        extraValue: ""
      }),
      Object.freeze({
        id: "sig_1794890241199",
        name: "President",
        fullName: "Mark Stemmans, P.E.",
        title: "President",
        company: "Convergent Method, LLC",
        email: "MarkSt@convergentmethod.com",
        phone: "(281) 979-1336",
        phoneLink: "2819791336",
        hasLogo: true,
        logo: "https://convergentmethod.com/media/CMFullColor.png",
        color: "#6fa0db",
        hasDisclaimer: true,
        hasEco: true,
        hasExtraDetail: false,
        extraLabel: "Dealer #:",
        extraValue: ""
      })
    ])
  }),
  "amandast@convergentmethod.com": Object.freeze({
    name: "Amanda Stemmans, MS, APRN, FNP-C",
    title: "Project Coordinator",
    company: "Convergent Method, LLC",
    email: "AmandaSt@convergentmethod.com",
    phone: "(281) 979-2162",
    phoneLink: "2819792162",
    hasLogo: true,
    hasDisclaimer: true,
    hasEco: true,
    hasExtraDetail: false,
    extraLabel: "Dealer #:",
    extraValue: "",
    signatures: Object.freeze([
      Object.freeze({
        id: "sig_1794890275640_1",
        name: "Project",
        fullName: "Amanda Stemmans, MS, APRN, FNP-C",
        title: "Project Coordinator",
        company: "Convergent Method, LLC",
        email: "AmandaSt@convergentmethod.com",
        phone: "(281) 979-2162",
        phoneLink: "2819792162",
        hasLogo: true,
        logo: "https://convergentmethod.com/media/CMFullColor.png",
        color: "#6fa0db",
        hasDisclaimer: true,
        hasEco: true,
        hasExtraDetail: false,
        extraLabel: "Dealer #:",
        extraValue: ""
      }),
      Object.freeze({
        id: "sig_1794891481602",
        name: "COO",
        fullName: "Amanda Stemmans, MS, APRN, FNP-C",
        title: "Chief Operations Officer",
        company: "Convergent Method, LLC",
        email: "AmandaSt@convergentmethod.com",
        phone: "(281) 979-2162",
        phoneLink: "2819792162",
        hasLogo: true,
        logo: "https://convergentmethod.com/media/CMFullColor.png",
        color: "#6fa0db",
        hasDisclaimer: true,
        hasEco: false,
        hasExtraDetail: false,
        extraLabel: "Dealer #:",
        extraValue: ""
      })
    ])
  })
});

// If true, forces Mark Stemmans' signature across all accounts regardless of logged-in user
var FORCE_MARK_SIGNATURE = false;

Office.onReady(function () {});

function onNewMessageComposeHandler(event) {
  applySignature(event);
}

function onNewAppointmentOrganizerHandler(event) {
  applySignature(event);
}

/**
 * Escapes characters to prevent HTML/XSS injection.
 */
function escapeHtml(text) {
  if (text === null || text === undefined) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Sanitizes phone numbers for tel: links.
 */
function sanitizePhoneLink(link) {
  if (!link) return "";
  return String(link).replace(/[^0-9+]/g, "");
}

function buildSignatureHtml(user) {
  var safeName = escapeHtml(user.name);
  var safeTitle = escapeHtml(user.title);
  var safeCompany = escapeHtml(user.company || "Convergent Method, LLC");
  var safeEmail = escapeHtml(user.email);
  var safePhone = escapeHtml(user.phone);
  var safePhoneLink = sanitizePhoneLink(user.phoneLink);
  var hasLogo = user.hasLogo !== false;
  var hasDisclaimer = user.hasDisclaimer !== false;
  var hasEco = user.hasEco !== false;

  var extraDetailHtml = '';
  if (user.hasExtraDetail && user.extraValue) {
    var safeExtraLabel = escapeHtml(user.extraLabel || 'Dealer #:');
    var safeExtraValue = escapeHtml(user.extraValue);
    extraDetailHtml = 
      '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 12px; line-height: 1.4; margin: 0; white-space: nowrap;">' +
        '<span style="font-weight: 600; color: #333333;">' + safeExtraLabel + '&nbsp;</span><span style="color: #444444;">' + safeExtraValue + '</span>' +
      '</div>';
  }

  var topSection = '';
  if (hasLogo) {
    topSection = 
      '<tr>' +
        '<td valign="middle" width="153" style="width: 153px; min-width: 153px; max-width: 153px; padding: 0 24px 0 0; border-right: 2px solid #6FA0DB; vertical-align: middle;">' +
          '<a href="https://convergentmethod.com" target="_blank" style="text-decoration: none;">' +
            '<img src="https://convergentmethod.com/media/CMFullColor.png" alt="Convergent Method Logo" width="145" border="0" style="display: block; width: 145px; max-width: 145px; height: auto; border: 0; outline: none; text-decoration: none;" />' +
          '</a>' +
        '</td>' +
        '<td valign="middle" width="387" style="width: 387px; min-width: 387px; padding: 0 0 0 16px; vertical-align: middle; white-space: nowrap;">' +
          '<div style="font-family: Century Gothic, CenturyGothic, Apple Gothic, Avenir Next, Avenir, Segoe UI, Arial, sans-serif; font-size: 17px; font-weight: 700; color: #6FA0DB; letter-spacing: 0.2px; line-height: 1.2; margin: 0 0 3px 0; white-space: nowrap;">' +
            safeName +
          '</div>' +
          '<div style="font-family: Century Gothic, CenturyGothic, Apple Gothic, Avenir Next, Avenir, Segoe UI, Arial, sans-serif; font-size: 13px; font-weight: 600; color: #222222; line-height: 1.3; margin: 0 0 2px 0; white-space: nowrap;">' +
            safeTitle +
          '</div>' +
          '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 13px; font-weight: 400; color: #555555; line-height: 1.3; margin: 0 0 8px 0; white-space: nowrap;">' +
            safeCompany +
          '</div>' +
          '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 12px; line-height: 1.4; margin: 0 0 2px 0; white-space: nowrap;">' +
            '<span style="font-weight: 600; color: #333333;">E:&nbsp;</span>' +
            '<a href="mailto:' + encodeURIComponent(user.email) + '" style="color: #6FA0DB; text-decoration: none; font-weight: 500;">' +
              '<span style="color: #6FA0DB; text-decoration: none; font-weight: 500;">' + safeEmail + '</span>' +
            '</a>' +
          '</div>' +
          '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 12px; line-height: 1.4; margin: 0; white-space: nowrap;">' +
            '<span style="font-weight: 600; color: #333333;">P:&nbsp;</span>' +
            '<a href="tel:' + safePhoneLink + '" style="color: #444444; text-decoration: none;">' +
              '<span style="color: #444444; text-decoration: none;">' + safePhone + '</span>' +
            '</a>' +
          '</div>' +
          extraDetailHtml +
        '</td>' +
      '</tr>';
  } else {
    topSection = 
      '<tr>' +
        '<td colspan="2" valign="top" style="padding: 0 0 6px 0; vertical-align: top;">' +
          '<div style="font-family: Century Gothic, CenturyGothic, Apple Gothic, Avenir Next, Avenir, Segoe UI, Arial, sans-serif; font-size: 17px; font-weight: 700; color: #6FA0DB; letter-spacing: 0.2px; line-height: 1.2; margin: 0 0 3px 0;">' +
            safeName +
          '</div>' +
          '<div style="font-family: Century Gothic, CenturyGothic, Apple Gothic, Avenir Next, Avenir, Segoe UI, Arial, sans-serif; font-size: 13px; font-weight: 600; color: #222222; line-height: 1.3; margin: 0 0 2px 0;">' +
            safeTitle +
          '</div>' +
          '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 13px; font-weight: 400; color: #555555; line-height: 1.3; margin: 0 0 6px 0;">' +
            safeCompany +
          '</div>' +
          '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 12px; line-height: 1.4; margin: 0 0 2px 0;">' +
            '<span style="font-weight: 600; color: #333333;">E:&nbsp;</span>' +
            '<a href="mailto:' + encodeURIComponent(user.email) + '" style="color: #6FA0DB; text-decoration: none; font-weight: 500;">' +
              '<span style="color: #6FA0DB; text-decoration: none; font-weight: 500;">' + safeEmail + '</span>' +
            '</a>' +
          '</div>' +
          '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 12px; line-height: 1.4; margin: 0;">' +
            '<span style="font-weight: 600; color: #333333;">P:&nbsp;</span>' +
            '<a href="tel:' + safePhoneLink + '" style="color: #444444; text-decoration: none;">' +
              '<span style="color: #444444; text-decoration: none;">' + safePhone + '</span>' +
            '</a>' +
          '</div>' +
          extraDetailHtml +
        '</td>' +
      '</tr>';
  }

  var bottomSection = '';
  if (hasEco || hasDisclaimer) {
    bottomSection = 
      '<tr><td colspan="2" height="14" style="height: 14px; font-size: 0; line-height: 0; padding: 0;">&nbsp;</td></tr>' +
      '<tr>' +
        '<td colspan="2" style="padding: 0; margin: 0; font-family: Aptos, Calibri, Arial, sans-serif;">' +
          (hasEco ? '<div style="font-size: 10px; line-height: 14px; color: #6FA0DB; font-style: italic; padding-bottom: 6px;">Please do not print this e-mail unless necessary</div>' : '') +
          (hasDisclaimer ? '<div style="font-size: 10px; line-height: 14px; color: #888888;"><strong style="color: #666666;">Notice:</strong> This message is confidential, is intended only for the named recipient(s) and may contain information that is privileged or exempt from disclosure under applicable law. If you are not the intended recipient(s), you are notified that the dissemination, distribution or copying of this message is strictly prohibited. If you received this message and are not an intended recipient, please delete it from your computer.</div>' : '') +
        '</td>' +
      '</tr>';
  }

  return (
    '<!-- Microsoft 365 Auto Signature - Convergent Method, LLC -->' +
    '<div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 13px; color: #333333; line-height: 1.4; padding-top: 16px;">' +
      '<table width="560" cellpadding="0" cellspacing="0" border="0" style="width: 560px; min-width: 560px; max-width: 560px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-family: Aptos, Calibri, Arial, sans-serif;">' +
        topSection +
        bottomSection +
      '</table>' +
    '</div>'
  );
    '<div style="font-family: \'Aptos\', Calibri, Arial, sans-serif; font-size: 13px; color: #333333; line-height: 1.4; padding-top: 16px;">' +
      '<table width="560" cellpadding="0" cellspacing="0" border="0" style="width: 560px; min-width: 560px; max-width: 560px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-family: \'Aptos\', Calibri, Arial, sans-serif;">' +
        topSection +
        bottomSection +
      '</table>' +
    '</div>'
  );
}

function resolveCurrentUserInfo() {
  if (FORCE_MARK_SIGNATURE) {
    return TEAM_DIRECTORY["markst@convergentmethod.com"] || TEAM_DIRECTORY[Object.keys(TEAM_DIRECTORY)[0]];
  }
  try {
    var profile = Office.context.mailbox.userProfile;
    var email = (profile && profile.emailAddress) ? profile.emailAddress.trim().toLowerCase() : "";
    var displayName = (profile && profile.displayName) ? profile.displayName.trim() : "";

    if (email && TEAM_DIRECTORY[email]) return TEAM_DIRECTORY[email];

    var nameLower = displayName.toLowerCase();
    for (var key in TEAM_DIRECTORY) {
      if (email.indexOf(key.split('@')[0]) !== -1 || (nameLower && nameLower.indexOf(key.split('@')[0]) !== -1)) {
        return TEAM_DIRECTORY[key];
      }
    }

    return null;
  } catch (e) {
    return null;
  }
}

function applySignature(event) {
  try {
    var item = Office.context.mailbox.item;
    if (item && item.body && typeof item.body.setSignatureAsync === "function") {
      var user = resolveCurrentUserInfo();
      if (!user) {
        event.completed({ allowEvent: true });
        return;
      }
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

if (typeof Office !== "undefined" && Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("insertSignatureButtonHandler", insertSignatureButtonHandler);
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentOrganizerHandler", onNewAppointmentOrganizerHandler);
}
