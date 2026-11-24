/**
 * ============================================================================
 * MANUSCRIPT AUDIT TOOLKIT — IN-PLACE MULTI-TAB GOOGLE DOCS & SHEETS UPDATER
 * ============================================================================
 * Repository: https://github.com/gf777/manuscript-audit-toolkit
 *
 * Solves two fundamental Google Workspace engineering hurdles simultaneously:
 *   1. Native Suggesting-Mode Edits & Phrase-Anchored Comments:
 *      Updates a live Google Doc IN PLACE (keeping the exact same URL, File ID,
 *      collaborator permissions, and Version History) from an audited `.docx`
 *      package containing OpenXML `<w:del>`/`<w:ins>` Track Changes and
 *      `<w:commentRangeStart>`/`<w:commentRangeEnd>` sidebar comments.
 *   2. Zero-Tab-Merge Multi-Tab Preservation (`updateGoogleDocPreservingAllTabs`):
 *      Because OpenXML `.docx` conversion flattens multi-tab Google Docs into a
 *      single tab, this script snapshots the live multi-tab hierarchy first via
 *      the Google Docs v1 REST API (`?includeTabsContent=true`), creates a full
 *      timestamped multi-tab backup in Google Drive, updates the target Main
 *      Manuscript tab in place from the single-tab audited `.docx`, restores its
 *      original tab title & emoji via `UpdateDocumentTabPropertiesRequest`, and
 *      automatically recreates and deep-clones all remaining tabs (including
 *      nested child sub-tabs, paragraphs, headings, lists, links, inline images,
 *      and tables) via `AddDocumentTabRequest` + `DocumentApp`!
 *
 * Setup (one-time per Apps Script project):
 *   1. Upload your `<manuscript>_AUDITED.docx` to Google Drive (keep as `.docx`).
 *   2. Open your live Google Doc -> Extensions -> Apps Script -> paste this file.
 *   3. In the left sidebar under "Services (+)", add BOTH:
 *        - Drive API
 *        - Google Docs API
 *   4. Click Run (`updateGoogleDocPreservingAllTabs` is selected by default).
 * ============================================================================
 */

var CONFIG = {
  // Leave empty ("") when running via Extensions -> Apps Script inside the target
  // Google Doc (auto-detected!), OR paste a Google Doc URL / File ID here:
  LIVE_MANUSCRIPT_DOC_ID: "",

  // Optional Tab ID of the Main Manuscript tab (e.g. "t.0" or from `?tab=t.xxxx`
  // in the URL). Leave empty ("") to default to the first tab:
  TARGET_MAIN_TAB_ID: "",

  // Leave empty ("") to auto-find `STAGED_AUDITED_DOCX_NAME` in Google Drive,
  // OR paste the uploaded `.docx` Drive URL / File ID here:
  STAGED_AUDITED_DOCX_ID: "",

  // Filename of the audited single-tab `.docx` uploaded to Google Drive:
  STAGED_AUDITED_DOCX_NAME: "audited_manuscript.docx",

  // Optional: Live Supplementary Tables Google Sheet URL / File ID & `.xlsx`:
  LIVE_SUPP_SHEET_ID: "",
  STAGED_AUDITED_XLSX_ID: "",
  STAGED_AUDITED_XLSX_NAME: "supplementary_tables_AUDITED_FIXED.xlsx"
};

/**
 * DEFAULT FUNCTION TO RUN:
 * Updates the live Google Doc IN PLACE with all native Suggesting-mode Track
 * Changes and phrase-anchored comments while preserving all other tabs unmerged!
 */
function updateGoogleDocPreservingAllTabs() {
  assertRequiredServices_();

  var targetDocId = resolveTargetDocId_();
  var explicitDocxId = extractDriveId_(CONFIG.STAGED_AUDITED_DOCX_ID);
  var sourceDocxFile = explicitDocxId
    ? DriveApp.getFileById(explicitDocxId)
    : findNewestDriveFileByName_(CONFIG.STAGED_AUDITED_DOCX_NAME);

  if (!sourceDocxFile) {
    throw new Error(
      "Could not find '" + CONFIG.STAGED_AUDITED_DOCX_NAME + "' in Google Drive. " +
      "Upload your audited .docx file to Drive (keep as .docx) or set CONFIG.STAGED_AUDITED_DOCX_ID."
    );
  }

  if (targetDocId === sourceDocxFile.getId()) {
    throw new Error("Target Google Doc ID and staged .docx ID are identical — aborting.");
  }

  var targetFile = DriveApp.getFileById(targetDocId);
  var docTitle = targetFile.getName();

  // --------------------------------------------------------------------------
  // PHASE 1: Snapshot live multi-tab hierarchy & save full backup in Drive
  // --------------------------------------------------------------------------
  Logger.log("=== PHASE 1: SNAPSHOTTING LIVE MULTI-TAB TREE & CREATING BACKUP ===");
  var origDocJson = docsRestGetWithTabs_(targetDocId);
  var flatOrigTabs = [];
  flattenJsonTabsPreOrder_(origDocJson.tabs || [], null, flatOrigTabs);
  Logger.log("Detected " + flatOrigTabs.length + " tab(s) in live Google Doc.");

  if (flatOrigTabs.length === 0) {
    throw new Error("Unexpected empty tabs array returned by Google Docs v1 API.");
  }

  var mainTabIdx = 0;
  if (CONFIG.TARGET_MAIN_TAB_ID) {
    for (var m = 0; m < flatOrigTabs.length; m++) {
      if (flatOrigTabs[m].origTabId === CONFIG.TARGET_MAIN_TAB_ID) {
        mainTabIdx = m;
        break;
      }
    }
  }
  var mainTabInfo = flatOrigTabs[mainTabIdx];

  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  var backupName = docTitle + " [PRE-AUDIT MULTI-TAB BACKUP " + timestamp + "]";
  var backupFile = targetFile.makeCopy(backupName);
  var backupDocId = backupFile.getId();
  Logger.log("Created safety backup: " + backupName + " (" + backupFile.getUrl() + ")");

  // --------------------------------------------------------------------------
  // PHASE 2: In-place update of Main Manuscript tab from audited .docx
  // --------------------------------------------------------------------------
  Logger.log("=== PHASE 2: IN-PLACE CONVERSION OF AUDITED .DOCX ONTO MAIN TAB ===");
  var docxBlob = sourceDocxFile.getBlob();
  docxBlob.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

  overwriteDriveFileInPlace_(targetDocId, docTitle, "application/vnd.google-apps.document", docxBlob);
  Utilities.sleep(1500);

  var afterConvertJson = docsRestGetWithTabs_(targetDocId);
  var convertedMainTabId = afterConvertJson.tabs[0].tabProperties.tabId;

  var mainPropsToRestore = {
    tabId: convertedMainTabId,
    title: mainTabInfo.title || "Manuscript"
  };
  var fieldMask = ["title"];
  if (mainTabInfo.iconEmoji) {
    mainPropsToRestore.iconEmoji = mainTabInfo.iconEmoji;
    fieldMask.push("iconEmoji");
  }
  docsRestBatchUpdate_(targetDocId, [
    {
      updateDocumentTabProperties: {
        tabProperties: mainPropsToRestore,
        fields: fieldMask.join(",")
      }
    }
  ]);

  var origToNewTabId = {};
  origToNewTabId[mainTabInfo.origTabId] = convertedMainTabId;

  // If single-tab document, we are already finished!
  if (flatOrigTabs.length === 1) {
    Logger.log("=== COMPLETE (SINGLE-TAB DOCUMENT UPDATED IN PLACE) ===");
    Logger.log("Live URL: https://docs.google.com/document/d/" + targetDocId + "/edit");
    return;
  }

  // --------------------------------------------------------------------------
  // PHASE 3: Recreate all remaining tabs (preserving order, icons, & nesting)
  // --------------------------------------------------------------------------
  Logger.log("=== PHASE 3: RECREATING REMAINING TABS UNMERGED IN SIDEBAR ===");
  for (var i = 0; i < flatOrigTabs.length; i++) {
    if (i === mainTabIdx) continue;
    var info = flatOrigTabs[i];
    var addProps = { title: info.title };

    if (info.parentOrigTabId && origToNewTabId[info.parentOrigTabId]) {
      addProps.parentTabId = origToNewTabId[info.parentOrigTabId];
    } else if (!info.parentOrigTabId && i < mainTabIdx) {
      addProps.index = info.index;
    }
    if (info.iconEmoji) {
      addProps.iconEmoji = info.iconEmoji;
    }

    var addResp = docsRestBatchUpdate_(targetDocId, [
      { addDocumentTab: { tabProperties: addProps } }
    ]);
    var createdTabId = addResp.replies[0].addDocumentTab.tabProperties.tabId;
    origToNewTabId[info.origTabId] = createdTabId;
    Logger.log("  Recreated tab [" + i + "] \"" + info.title + "\" -> " + createdTabId);
  }

  Utilities.sleep(1000);

  // --------------------------------------------------------------------------
  // PHASE 4: Deep-copy content, images, lists, & tables from backup tabs
  // --------------------------------------------------------------------------
  Logger.log("=== PHASE 4: POPULATING RECREATED TABS FROM MULTI-TAB BACKUP ===");
  var backupDoc = DocumentApp.openById(backupDocId);
  var liveDoc = DocumentApp.openById(targetDocId);

  var flatBackupDocTabs = [];
  collectDocAppTabsPreOrder_(backupDoc.getTabs(), flatBackupDocTabs);

  for (var k = 0; k < flatOrigTabs.length; k++) {
    if (k === mainTabIdx) continue;
    var tabMeta = flatOrigTabs[k];
    var dstTabId = origToNewTabId[tabMeta.origTabId];

    var srcTabObj = backupDoc.getTab(tabMeta.origTabId) || flatBackupDocTabs[k];
    var dstTabObj = liveDoc.getTab(dstTabId);
    if (!srcTabObj || !dstTabObj) continue;

    var copiedCount = copyTabBodyElements_(
      srcTabObj.asDocumentTab().getBody(),
      dstTabObj.asDocumentTab().getBody()
    );
    Logger.log("  Populated tab \"" + tabMeta.title + "\" (" + copiedCount + " elements)");
  }

  liveDoc.saveAndClose();
  Logger.log("=== COMPLETE: ALL TABS PRESERVED UNMERGED IN LIVE GOOGLE DOC! ===");
  Logger.log("Live URL: https://docs.google.com/document/d/" + targetDocId + "/edit");
  Logger.log("Backup URL: " + backupFile.getUrl());
}

/**
 * Optional Helper: Updates a live Google Sheet IN PLACE from an audited `.xlsx`
 * workbook while preserving its Google Sheet URL, permissions, & Version History.
 */
function updateGoogleSheetInPlace() {
  if (typeof Drive === "undefined" || !Drive.Files) {
    throw new Error("Please add 'Drive API' under Services (+) first.");
  }
  var targetSheetId = extractDriveId_(CONFIG.LIVE_SUPP_SHEET_ID);
  if (!targetSheetId) {
    throw new Error("Set CONFIG.LIVE_SUPP_SHEET_ID to your target Google Sheet URL or ID.");
  }
  var explicitXlsxId = extractDriveId_(CONFIG.STAGED_AUDITED_XLSX_ID);
  var sourceXlsx = explicitXlsxId
    ? DriveApp.getFileById(explicitXlsxId)
    : findNewestDriveFileByName_(CONFIG.STAGED_AUDITED_XLSX_NAME);
  if (!sourceXlsx) {
    throw new Error("Could not find '" + CONFIG.STAGED_AUDITED_XLSX_NAME + "' in Google Drive.");
  }

  var targetFile = DriveApp.getFileById(targetSheetId);
  var sheetTitle = targetFile.getName();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  var backup = targetFile.makeCopy(sheetTitle + " [PRE-AUDIT BACKUP " + timestamp + "]");
  Logger.log("Created Sheet backup: " + backup.getUrl());

  var xlsxBlob = sourceXlsx.getBlob();
  xlsxBlob.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  overwriteDriveFileInPlace_(targetSheetId, sheetTitle, "application/vnd.google-apps.spreadsheet", xlsxBlob);
  Logger.log("Updated Google Sheet in place: https://docs.google.com/spreadsheets/d/" + targetSheetId + "/edit");
}

/**
 * Optional Helper: Audits all historical, resolved, and deleted comment threads
 * on the target Google Doc via Drive Comments API.
 */
function inspectAllHistoricalCommentsIncludingDeleted() {
  var fileId = resolveTargetDocId_();
  var pageToken = null;
  var total = 0, active = 0, resolved = 0, deleted = 0;
  do {
    var resp;
    try {
      var optsV3 = { includeDeleted: true, pageSize: 100, fields: "*" };
      if (pageToken) optsV3.pageToken = pageToken;
      resp = Drive.Comments.list(fileId, optsV3);
    } catch (e) {
      var optsV2 = { includeDeleted: true, maxResults: 100, fields: "*" };
      if (pageToken) optsV2.pageToken = pageToken;
      resp = Drive.Comments.list(fileId, optsV2);
    }
    var items = resp.comments || resp.items || [];
    for (var i = 0; i < items.length; i++) {
      var c = items[i];
      total++;
      var st = c.deleted ? "DELETED" : (c.status === "resolved" || c.resolved ? "RESOLVED" : "ACTIVE");
      if (st === "DELETED") deleted++;
      else if (st === "RESOLVED") resolved++;
      else active++;
      var author = (c.author && c.author.displayName) ? c.author.displayName : "Unknown";
      var anchor = (c.quotedFileContent && c.quotedFileContent.value) ? c.quotedFileContent.value : "";
      Logger.log("[" + total + "] " + st + " | " + author + " | Anchor=\"" + anchor + "\" | " + (c.content || ""));
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);
  Logger.log("Total=" + total + " | Active=" + active + " | Resolved=" + resolved + " | Deleted=" + deleted);
}

// ============================================================================
// INTERNAL HELPERS
// ============================================================================

function copyTabBodyElements_(srcBody, dstBody) {
  var numChildren = srcBody.getNumChildren();
  var appendedCount = 0;

  for (var c = 0; c < numChildren; c++) {
    var elem = srcBody.getChild(c);
    var type = elem.getType();
    try {
      if (type === DocumentApp.ElementType.PARAGRAPH) {
        dstBody.appendParagraph(elem.asParagraph().copy());
        appendedCount++;
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        var srcLi = elem.asListItem();
        var newLi = dstBody.appendListItem(srcLi.copy());
        try {
          newLi.setGlyphType(srcLi.getGlyphType());
        } catch (eGlyph) {}
        appendedCount++;
      } else if (type === DocumentApp.ElementType.TABLE) {
        dstBody.appendTable(elem.asTable().copy());
        appendedCount++;
      } else if (type === DocumentApp.ElementType.PAGE_BREAK) {
        dstBody.appendPageBreak(elem.asPageBreak().copy());
        appendedCount++;
      } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
        dstBody.appendHorizontalRule();
        appendedCount++;
      }
    } catch (eElem) {
      try {
        dstBody.appendParagraph(elem.asText().getText());
        appendedCount++;
      } catch (eInner) {}
    }
  }

  if (appendedCount > 0 && dstBody.getNumChildren() > 1) {
    try {
      dstBody.removeChild(dstBody.getChild(0));
    } catch (eRm) {}
  }
  return appendedCount;
}

function flattenJsonTabsPreOrder_(tabsArray, parentOrigTabId, outList) {
  for (var i = 0; i < tabsArray.length; i++) {
    var t = tabsArray[i];
    var p = t.tabProperties || {};
    outList.push({
      origTabId: p.tabId,
      title: p.title || "",
      iconEmoji: p.iconEmoji || "",
      index: typeof p.index === "number" ? p.index : i,
      parentOrigTabId: parentOrigTabId
    });
    if (t.childTabs && t.childTabs.length > 0) {
      flattenJsonTabsPreOrder_(t.childTabs, p.tabId, outList);
    }
  }
}

function collectDocAppTabsPreOrder_(docAppTabs, outList) {
  for (var i = 0; i < docAppTabs.length; i++) {
    var tab = docAppTabs[i];
    outList.push(tab);
    var children = tab.getChildTabs();
    if (children && children.length > 0) {
      collectDocAppTabsPreOrder_(children, outList);
    }
  }
}

function docsRestGetWithTabs_(docId) {
  var url = "https://docs.googleapis.com/v1/documents/" + encodeURIComponent(docId) + "?includeTabsContent=true";
  var resp = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    throw new Error("Docs API GET failed (" + resp.getResponseCode() + "): " + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

function docsRestBatchUpdate_(docId, requestsArray) {
  var url = "https://docs.googleapis.com/v1/documents/" + encodeURIComponent(docId) + ":batchUpdate";
  var resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ requests: requestsArray }),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    throw new Error("Docs API batchUpdate failed (" + resp.getResponseCode() + "): " + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

function overwriteDriveFileInPlace_(targetId, fileName, targetMimeType, mediaBlob) {
  try {
    Drive.Files.update(
      { name: fileName, mimeType: targetMimeType },
      targetId,
      mediaBlob,
      { supportsAllDrives: true }
    );
  } catch (eV3Mime) {
    try {
      Drive.Files.update(
        { name: fileName },
        targetId,
        mediaBlob,
        { supportsAllDrives: true }
      );
    } catch (eV3) {
      Drive.Files.update(
        { title: fileName, mimeType: targetMimeType },
        targetId,
        mediaBlob,
        { supportsAllDrives: true, convert: true, newRevision: true }
      );
    }
  }
}

function resolveTargetDocId_() {
  var fromConfig = extractDriveId_(CONFIG.LIVE_MANUSCRIPT_DOC_ID);
  if (fromConfig) return fromConfig;
  try {
    var activeDoc = DocumentApp.getActiveDocument();
    if (activeDoc) return activeDoc.getId();
  } catch (e) {}
  throw new Error(
    "Could not auto-detect target Google Doc. Either open this script from inside the " +
    "Google Doc (Extensions -> Apps Script) OR set CONFIG.LIVE_MANUSCRIPT_DOC_ID."
  );
}

function assertRequiredServices_() {
  var missing = [];
  if (typeof Drive === "undefined" || !Drive.Files) missing.push("'Drive API'");
  if (typeof Docs === "undefined" || !Docs.Documents) missing.push("'Google Docs API'");
  if (missing.length > 0) {
    throw new Error(
      "Missing Advanced Service(s): " + missing.join(" and ") + "! " +
      "In the left sidebar of Apps Script, click 'Services (+)', add " +
      missing.join(" and ") + ", then click Run again."
    );
  }
}

function extractDriveId_(input) {
  if (!input) return "";
  var trimmed = String(input).trim();
  var match = trimmed.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (match && match[1]) return match[1];
  var idParam = trimmed.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (idParam && idParam[1]) return idParam[1];
  return trimmed;
}

function findNewestDriveFileByName_(fileName) {
  var files = DriveApp.getFilesByName(fileName);
  var newestFile = null;
  var newestTime = 0;
  while (files.hasNext()) {
    var f = files.next();
    if (f.isTrashed()) continue;
    var t = f.getLastUpdated().getTime();
    if (t > newestTime) {
      newestTime = t;
      newestFile = f;
    }
  }
  return newestFile;
}
