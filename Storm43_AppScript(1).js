// Storm43 MIS — Google Apps Script v5.3
// 1. Paste this entire file replacing all existing code
// 2. Change STUDIO_EMAIL to your real email address
// 3. Save (Ctrl+S)
// 4. Run testEmail to verify emails work
// 5. Run testScript to verify chunk storage works
// 6. Deploy > Manage deployments > Edit > New version > Deploy

var STUDIO_EMAIL = 'design@storm43.co.za'; // <-- CHANGE THIS

// ── doGet ──
function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};

  try {
    if (action === 'initFileUpload') {
      var jobId = e.parameter.jobId;
      var totalChunks = parseInt(e.parameter.totalChunks || '1');
      var fileType = e.parameter.fileType || 'pdf';
      var fileName = e.parameter.fileName || 'design.pdf';
      var sheetName = 'f_' + jobId;
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        sheet.clearContents();
      } else {
        sheet = ss.insertSheet(sheetName);
      }
      sheet.getRange(1, 1).setValue(JSON.stringify({
        jobId: jobId, fileName: fileName, fileType: fileType,
        totalChunks: totalChunks, uploadedAt: new Date().toISOString()
      }));
      result = { success: true, jobId: jobId, totalChunks: totalChunks };
    } else if (action === 'verifyChunk') {
      var jobId = e.parameter.jobId;
      var chunkIndex = parseInt(e.parameter.chunkIndex || '0');
      var sheetName = 'f_' + jobId;
      var sheet = ss.getSheetByName(sheetName);
      var exists = false;
      if (sheet) {
        var val = sheet.getRange(chunkIndex + 2, 1).getValue();
        exists = val && val.length > 0;
      }
      result = { exists: exists, chunkIndex: chunkIndex };
    } else if (action === 'read') {
      result = readAllData(ss);
    } else if (action === 'write') {
      var payload = e.parameter.payload;
      if (payload) {
        var data = JSON.parse(decodeURIComponent(payload));
        writeAllData(ss, data);
        result = { success: true, timestamp: new Date().toISOString() };
      }
    } else if (action === 'saveFileChunk') {
      result = saveFileChunkFromGet(ss, e.parameter);
    } else if (action === 'getFile') {
      // Try Drive first, fall back to Sheets chunks
      var driveResult = getFileFromDrive(e.parameter.jobId);
      if (driveResult.hasFile) {
        result = driveResult;
      } else {
        result = getFile(ss, e.parameter.jobId);
      }
    } else if (action === 'deleteFile') {
      result = deleteFile(ss, e.parameter.jobId);
    } else if (action === 'sendEmail') {
      Logger.log('sendEmail action received');
      var emailResult = sendNotificationEmail(e.parameter);
      result = emailResult || { success: true };
    } else if (action === 'sendAnnotationReport') {
      Logger.log('sendAnnotationReport action received');
      var reportResult = sendAnnotationReportEmail(e.parameter);
      result = reportResult || { success: true };
    } else if (action === 'readBriefs') {
      // ── NEW: Read all submitted briefs ──
      result = readBriefs(ss);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    Logger.log('doGet error: ' + err.message);
    result = { error: err.message };
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── doPost ──
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'saveFileChunk') {
      var chunkResult = saveFileChunk(ss, data.jobId, parseInt(data.chunkIndex),
        parseInt(data.totalChunks), data.chunk, data.fileType, data.fileName);
      return respondPost(chunkResult);
    }

    if (data.action === 'saveLead') {
      var leadResult = saveLead(ss, data.lead);
      return respondPost(leadResult);
    }

    if (data.action === 'saveToDrive') {
      Logger.log('saveToDrive for jobId: ' + data.jobId);
      var driveResult = saveToDrive(data.jobId, data.fileName, data.fileType, data.dataUrl);
      return respondPost(driveResult);
    }

    if (data.action === 'getFile') {
      Logger.log('getFile via POST for jobId: ' + data.jobId);
      var driveFile = getFileFromDrive(data.jobId);
      if (driveFile.hasFile) return respondPost(driveFile);
      var fileResult = getFile(ss, data.jobId);
      return respondPost(fileResult);
    }

    if (data.action === 'sendEmail') {
      Logger.log('sendEmail via POST received');
      var emailResult = sendNotificationEmail(data);
      return respondPost(emailResult || { success: true });
    }

    if (data.action === 'sendAnnotationReport') {
      Logger.log('sendAnnotationReport via POST received');
      var reportResult = sendAnnotationReportEmail(data);
      return respondPost(reportResult || { success: true });
    }

    if (data.action === 'sendAssignmentEmail') {
      Logger.log('sendAssignmentEmail via POST received for: ' + data.designerEmail);
      var assignResult = sendDesignerAssignmentEmail(data);
      return respondPost(assignResult || { success: true });
    }

    // ── NEW: Client brief submission ──
    if (data.action === 'writeBrief') {
      var briefResult = writeBriefData(ss, data.brief);
      // Send notification email to studio
      sendBriefNotificationEmail(data.brief);
      return respondPost(briefResult);
    }

    writeAllData(ss, data);
    return respondPost({ success: true, timestamp: new Date().toISOString() });

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return respondPost({ error: err.message });
  }
}

function respondPost(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET HELPERS ──
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1).setValue('json_data');
    sheet.getRange(2, 1).setValue('[]');
  }
  return sheet;
}

// ── READ ALL DATA ──
function readAllData(ss) {
  var keys = ['jobs','clients','tasks','timeEntries','costItems',
              'designFiles','users','sections','approvals','driveFiles'];
  var result = {};

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var sheet = getOrCreateSheet(ss, key);
    var raw = sheet.getRange(2, 1).getValue();
    try {
      result[key] = (raw && raw !== '') ? JSON.parse(raw) : [];
    } catch (e) {
      result[key] = [];
    }
  }

  var metaSheet = getOrCreateSheet(ss, 'meta');
  var metaRaw = metaSheet.getRange(2, 1).getValue();
  try {
    var meta = metaRaw ? JSON.parse(metaRaw) : {};
    result.jobCounter = meta.jobCounter || 1;
  } catch (e) {
    result.jobCounter = 1;
  }

  var fmSheet = ss.getSheetByName('fileManifest');
  result.fileManifest = [];
  if (fmSheet) {
    try {
      var raw2 = fmSheet.getRange(2, 1).getValue();
      result.fileManifest = raw2 ? JSON.parse(raw2) : [];
    } catch (e) {}
  }

  return result;
}

// ── WRITE ALL DATA ──
function writeAllData(ss, data) {
  // Partial update from portal — only approvals + job status fields
  if (data._partialUpdate) {
    Logger.log('Partial update received');

    if (data.approvals !== undefined && data.approvals.length > 0) {
      var appSheet = getOrCreateSheet(ss, 'approvals');
      // MERGE: read existing approvals and append new ones (don't overwrite)
      var existingApprovals = [];
      try {
        var existingRaw = appSheet.getRange(2, 1).getValue();
        existingApprovals = existingRaw ? JSON.parse(existingRaw) : [];
      } catch(e) { existingApprovals = []; }

      // Add new reviews that don't already exist (check by reviewId)
      var newApprovals = data.approvals || [];
      for (var ai = 0; ai < newApprovals.length; ai++) {
        var newRev = newApprovals[ai];
        var alreadyExists = false;
        for (var ei = 0; ei < existingApprovals.length; ei++) {
          if (existingApprovals[ei].reviewId === newRev.reviewId) {
            alreadyExists = true;
            break;
          }
        }
        if (!alreadyExists) existingApprovals.push(newRev);
      }

      appSheet.getRange(2, 1).setValue(JSON.stringify(existingApprovals));
      Logger.log('Approvals merged — total: ' + existingApprovals.length + ' (added: ' + newApprovals.length + ')');
    }

    if (data.jobs && data.jobs.length > 0) {
      var jobsSheet = getOrCreateSheet(ss, 'jobs');
      var existingJobs = [];
      try {
        var raw = jobsSheet.getRange(2, 1).getValue();
        existingJobs = raw ? JSON.parse(raw) : [];
      } catch (e) {}

      for (var i = 0; i < data.jobs.length; i++) {
        var update = data.jobs[i];
        for (var j = 0; j < existingJobs.length; j++) {
          if (existingJobs[j].id === update.id) {
            existingJobs[j].clientApproval  = update.clientApproval;
            existingJobs[j].lastReviewer    = update.lastReviewer;
            existingJobs[j].lastReviewDate  = update.lastReviewDate;
            existingJobs[j].reviewIteration = update.reviewIteration;
            break;
          }
        }
      }
      jobsSheet.getRange(2, 1).setValue(JSON.stringify(existingJobs));
      Logger.log('Jobs updated: ' + data.jobs.length);
    }
    return;
  }

  // Full update from MIS
  var keys = ['jobs','clients','tasks','timeEntries','costItems',
              'designFiles','users','sections','approvals','driveFiles'];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (data[key] === undefined) continue;
    var sheet = getOrCreateSheet(ss, key);
    if (key === 'designFiles') {
      var stripped = (data[key] || []).map(function(f) {
        var copy = {};
        for (var k in f) { if (k !== 'dataUrl') copy[k] = f[k]; }
        return copy;
      });
      sheet.getRange(2, 1).setValue(JSON.stringify(stripped));
    } else {
      sheet.getRange(2, 1).setValue(JSON.stringify(data[key] || []));
    }
  }

  var metaSheet = getOrCreateSheet(ss, 'meta');
  metaSheet.getRange(2, 1).setValue(JSON.stringify({
    jobCounter: data.jobCounter || 1,
    lastSaved: new Date().toISOString()
  }));
}

// ── FILE CHUNK STORAGE ──
function getFileSheetName(jobId) {
  return 'f_' + String(jobId).replace(/[^a-zA-Z0-9_]/g, '').slice(-12);
}

function saveFileChunkFromGet(ss, params) {
  return saveFileChunk(ss,
    params.jobId,
    parseInt(params.chunkIndex),
    parseInt(params.totalChunks),
    params.chunk,
    params.fileType || 'pdf',
    params.fileName || 'design.pdf'
  );
}

function saveFileChunk(ss, jobId, chunkIndex, totalChunks, chunk, fileType, fileName) {
  if (!jobId || chunk === undefined || chunk === null) {
    return { error: 'Missing params' };
  }

  var sheetName = getFileSheetName(jobId);
  var sheet = ss.getSheetByName(sheetName);

  if (chunkIndex === 0) {
    if (sheet) {
      sheet.clearContents();
    } else {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.getRange(1, 1).setValue(JSON.stringify({
      jobId: jobId,
      fileName: fileName,
      fileType: fileType,
      totalChunks: totalChunks,
      uploadedAt: new Date().toISOString()
    }));
  } else {
    if (!sheet) return { error: 'Sheet not found. Send chunk 0 first.' };
  }

  sheet.getRange(chunkIndex + 2, 1).setValue(chunk);

  var isComplete = (chunkIndex === totalChunks - 1);
  if (isComplete) {
    updateFileManifest(ss, jobId, sheetName, fileName, fileType, totalChunks);
  }

  return { success: true, complete: isComplete, chunkIndex: chunkIndex, totalChunks: totalChunks };
}

function updateFileManifest(ss, jobId, sheetName, fileName, fileType, totalChunks) {
  var sheet = ss.getSheetByName('fileManifest');
  if (!sheet) {
    sheet = ss.insertSheet('fileManifest');
    sheet.getRange(1, 1).setValue('manifest');
    sheet.getRange(2, 1).setValue('[]');
  }
  var manifest = [];
  try {
    var raw = sheet.getRange(2, 1).getValue();
    manifest = raw ? JSON.parse(raw) : [];
  } catch (e) {}
  manifest = manifest.filter(function(m) { return m.jobId !== jobId; });
  manifest.push({ jobId: jobId, sheetName: sheetName, fileName: fileName,
    fileType: fileType, totalChunks: totalChunks, uploadedAt: new Date().toISOString() });
  sheet.getRange(2, 1).setValue(JSON.stringify(manifest));
}

function getFile(ss, jobId) {
  if (!jobId) return { error: 'Missing jobId', hasFile: false };
  var sheetName = getFileSheetName(jobId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'No file stored for this project', hasFile: false };

  var meta = {};
  try {
    meta = JSON.parse(sheet.getRange(1, 1).getValue());
  } catch (e) {
    return { error: 'Bad file metadata', hasFile: false };
  }

  var totalChunks = meta.totalChunks || 1;
  var chunks = [];
  for (var i = 0; i < totalChunks; i++) {
    var chunk = sheet.getRange(i + 2, 1).getValue();
    if (!chunk) return { error: 'Missing chunk ' + i, hasFile: false };
    chunks.push(chunk);
  }

  return {
    success: true, hasFile: true, jobId: jobId,
    fileName: meta.fileName, fileType: meta.fileType,
    totalChunks: totalChunks, uploadedAt: meta.uploadedAt,
    dataUrl: chunks.join('')
  };
}

function deleteFile(ss, jobId) {
  try {
    var sheet = ss.getSheetByName(getFileSheetName(jobId));
    if (sheet) ss.deleteSheet(sheet);
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
// ── BRIEF PORTAL — NEW FUNCTIONS ──
// ════════════════════════════════════════════════════════════

// Write a submitted client brief to the 'briefs' sheet
function writeBriefData(ss, brief) {
  if (!brief) return { error: 'No brief data received' };

  var sheet = ss.getSheetByName('briefs');
  if (!sheet) {
    sheet = ss.insertSheet('briefs');
    // Create header row
    sheet.getRange(1, 1, 1, 17).setValues([[
      'Ref', 'Submitted', 'Status',
      'Client Name', 'Company', 'Email', 'Phone', 'Industry', 'Referral',
      'Project Name', 'Design Type', 'Project Type', 'Priority', 'Due Date', 'Budget',
      'Summary', 'Audience'
    ]]);
    // Format header row
    sheet.getRange(1, 1, 1, 17).setBackground('#0e0e0d').setFontColor('#5AB544').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Also store full JSON in a separate details column for easy retrieval
  var fullJson = JSON.stringify(brief);

  sheet.appendRow([
    brief.ref || '',
    brief.submittedAt ? new Date(brief.submittedAt).toLocaleString('en-ZA') : new Date().toLocaleString('en-ZA'),
    'New Brief',
    brief.clientName || '',
    brief.clientCompany || '',
    brief.clientEmail || '',
    brief.clientPhone || '',
    brief.clientIndustry || '',
    brief.referral || '',
    brief.projectName || '',
    brief.designType || '',
    brief.projectType || '',
    brief.priority || '',
    brief.dueDate || '',
    brief.budget || '',
    brief.summary || '',
    brief.audience || ''
  ]);

  // Store the full JSON details in a separate 'briefDetails' sheet for complete retrieval
  var detailSheet = ss.getSheetByName('briefDetails');
  if (!detailSheet) {
    detailSheet = ss.insertSheet('briefDetails');
    detailSheet.getRange(1, 1).setValue('ref');
    detailSheet.getRange(1, 2).setValue('json_data');
  }

  var lastRow = Math.max(detailSheet.getLastRow(), 1) + 1;
  detailSheet.getRange(lastRow, 1).setValue(brief.ref || '');
  detailSheet.getRange(lastRow, 2).setValue(fullJson);

  Logger.log('Brief saved: ' + (brief.ref || 'no-ref'));
  return { success: true, ref: brief.ref };
}

// Read all submitted briefs (for MIS inbox)
function readBriefs(ss) {
  var briefs = [];
  var detailSheet = ss.getSheetByName('briefDetails');
  if (!detailSheet) return { briefs: [] };

  var lastRow = detailSheet.getLastRow();
  if (lastRow < 2) return { briefs: [] };

  for (var i = 2; i <= lastRow; i++) {
    var jsonVal = detailSheet.getRange(i, 2).getValue();
    if (jsonVal) {
      try {
        briefs.push(JSON.parse(jsonVal));
      } catch (e) {
        Logger.log('Failed to parse brief row ' + i + ': ' + e.message);
      }
    }
  }

  return { briefs: briefs, count: briefs.length };
}

// Send email notification to studio when a new brief is submitted
function sendBriefNotificationEmail(brief) {
  if (!brief) return;
  Logger.log('Sending brief notification email to: ' + STUDIO_EMAIL);

  try {
    var subject = 'New Brief — ' + (brief.ref || '') + ' — ' + (brief.projectName || 'Unnamed Project');

    // ── HTML helper functions (defined at outer scope for Apps Script compatibility) ──
    // These are defined below as module-level functions: briefRow() and briefSection()

    // Build AI conversation block
    var aiBlock = '';
    if (brief.aiConversationSummary) {
      var lines = brief.aiConversationSummary.split('\n---\n');
      var aiRows = '';
      for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (!line) continue;
        var isUser = line.indexOf('USER:') === 0;
        aiRows += '<tr><td style="padding:8px 12px;vertical-align:top;width:80px">' +
          '<span style="font-size:10px;font-weight:700;font-family:monospace;color:' + (isUser ? '#1456a0' : '#5AB544') + '">' +
          (isUser ? 'CLIENT' : 'AI') + '</span></td>' +
          '<td style="padding:8px 12px;font-size:12px;color:#18180f;line-height:1.5;border-bottom:1px solid #f0efea">' +
          line.replace(/^USER:\s*/, '').replace(/^ASSISTANT:\s*/, '') + '</td></tr>';
      }
      if (aiRows) {
        aiBlock = '<div style="margin-bottom:20px">' +
          '<div style="background:#0f1f0f;padding:10px 16px;border-radius:8px 8px 0 0">' +
          '<span style="color:#5AB544;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;font-family:monospace">&#10022; AI Brief Conversation</span>' +
          '</div>' +
          '<table style="width:100%;border-collapse:collapse;background:#ffffff;border:1px solid #e0dfd9;border-top:none;border-radius:0 0 8px 8px">' +
          aiRows + '</table></div>';
      }
    }

    var htmlBody =
      '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f4f3ef;font-family:Inter,system-ui,sans-serif">' +
      '<div style="max-width:640px;margin:0 auto;padding:24px 16px">' +

      // Header
      '<div style="background:#0e0e0d;border-radius:12px 12px 0 0;padding:20px 24px;display:flex;justify-content:space-between;align-items:center;border-bottom:3px solid #5AB544">' +
      '<div>' +
      '<div style="font-size:16px;font-weight:700;color:#ffffff;letter-spacing:-0.3px">Storm43 MIS</div>' +
      '<div style="font-size:11px;color:rgba(255,255,255,0.35);font-family:monospace;letter-spacing:1px;margin-top:2px">NEW CLIENT BRIEF</div>' +
      '</div>' +
      '<div style="text-align:right">' +
      '<div style="font-size:16px;font-weight:700;color:#5AB544;font-family:monospace">' + (brief.ref || '') + '</div>' +
      '<div style="font-size:11px;color:rgba(255,255,255,0.3);font-family:monospace;margin-top:2px">' + new Date().toLocaleDateString('en-ZA', {day:'2-digit',month:'long',year:'numeric'}) + '</div>' +
      '</div></div>' +

      // Alert banner
      '<div style="background:#eaf5e6;border:1px solid #c8e8c0;border-top:none;padding:14px 20px;border-radius:0 0 8px 8px;margin-bottom:20px">' +
      '<div style="font-size:13px;color:#3B6D11;font-weight:600">&#128203; A new project brief has been submitted via the client portal.</div>' +
      '<div style="font-size:12px;color:#3B6D11;margin-top:4px;opacity:0.8">Log into your MIS to review and action this brief.</div>' +
      '</div>' +

      // Client Details
      briefSection('Client Details', '&#128100;',
        briefRow('Name', brief.clientName) +
        briefRow('Company', brief.clientCompany) +
        briefRow('Email', '<a href="mailto:' + (brief.clientEmail||'') + '" style="color:#1456a0">' + (brief.clientEmail||'') + '</a>') +
        briefRow('Phone', brief.clientPhone) +
        briefRow('Industry', brief.clientIndustry) +
        briefRow('Referral', brief.referral)
      ) +

      // Project Info
      briefSection('Project Information', '&#128203;',
        briefRow('Project Name', '<strong>' + (brief.projectName||'') + '</strong>') +
        briefRow('Design Type', brief.designType) +
        briefRow('Project Type', brief.projectType) +
        briefRow('Priority', brief.priority) +
        briefRow('Due Date', brief.dueDate || 'Not specified') +
        briefRow('Budget', brief.budget || 'Not specified')
      ) +

      // Brief
      '<div style="margin-bottom:20px">' +
      '<div style="background:#0e0e0d;padding:10px 16px;border-radius:8px 8px 0 0">' +
      '<span style="color:#5AB544;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;font-family:monospace">&#10024; Design Brief</span>' +
      '</div>' +
      '<div style="background:#ffffff;border:1px solid #e0dfd9;border-top:none;border-radius:0 0 8px 8px;padding:0">' +

      (brief.summary ? '<div style="padding:14px 16px;border-bottom:1px solid #f0efea"><div style="font-size:10px;font-family:monospace;color:#6b6b60;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:6px">Summary</div><div style="font-size:13px;color:#18180f;line-height:1.7;padding:10px 12px;background:#f4f3ef;border-radius:6px;border-left:3px solid #5AB544">' + brief.summary + '</div></div>' : '') +

      (brief.audience ? '<div style="padding:14px 16px;border-bottom:1px solid #f0efea"><div style="font-size:10px;font-family:monospace;color:#6b6b60;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:4px">Target Audience</div><div style="font-size:13px;color:#18180f">' + brief.audience + '</div></div>' : '') +

      (brief.tone ? '<div style="padding:14px 16px;border-bottom:1px solid #f0efea"><div style="font-size:10px;font-family:monospace;color:#6b6b60;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:4px">Brand Tone</div><div style="font-size:13px;color:#18180f">' + brief.tone + '</div></div>' : '') +

      (brief.colours ? '<div style="padding:14px 16px;border-bottom:1px solid #f0efea"><div style="font-size:10px;font-family:monospace;color:#6b6b60;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:4px">Colour Preferences</div><div style="font-size:13px;color:#18180f">' + brief.colours + '</div></div>' : '') +

      (brief.competitors ? '<div style="padding:14px 16px;border-bottom:1px solid #f0efea"><div style="font-size:10px;font-family:monospace;color:#6b6b60;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:4px">Competitors / Inspiration</div><div style="font-size:13px;color:#18180f">' + brief.competitors + '</div></div>' : '') +

      (brief.notes ? '<div style="padding:14px 16px"><div style="font-size:10px;font-family:monospace;color:#6b6b60;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:4px">Additional Notes</div><div style="font-size:13px;color:#18180f">' + brief.notes + '</div></div>' : '') +

      '</div></div>' +

      // Files
      (brief.attachedCount && brief.attachedCount > 0 ?
        briefSection('Attached Files', '&#128206;',
          briefRow('Files uploaded', brief.attachedCount + ' file(s) submitted by client')
        ) : '') +

      // AI Conversation
      aiBlock +

      // CTA Button
      '<div style="text-align:center;margin:24px 0">' +
      '<a href="https://storm43design.github.io/Storm43-MIS/" style="display:inline-block;background:#5AB544;color:#ffffff;text-decoration:none;padding:13px 28px;border-radius:8px;font-size:14px;font-weight:700;letter-spacing:0.3px">Open Storm43 MIS &#8594;</a>' +
      '</div>' +

      // Footer
      '<div style="text-align:center;padding:16px;border-top:1px solid #e0dfd9">' +
      '<div style="font-size:11px;color:#a8a89a;font-family:monospace">Storm43 Creative Media Agency &bull; ' + (brief.ref || '') + '</div>' +
      '</div>' +

      '</div></body></html>';

    MailApp.sendEmail({
      to: STUDIO_EMAIL,
      subject: subject,
      htmlBody: htmlBody
    });

    Logger.log('Brief HTML notification email sent successfully');

    // ── CLIENT ACKNOWLEDGEMENT ──
    Logger.log('Client email field received: [' + (brief.clientEmail || 'EMPTY') + ']');
    var clientEmailAddr = (brief.clientEmail || '').trim();
    if (clientEmailAddr.length > 4 && clientEmailAddr.indexOf('@') > 0) {
      var clientSubject = 'Brief Received — ' + (brief.ref || '') + ' — Storm43';
      Logger.log('Attempting to send client acknowledgement to: ' + clientEmailAddr);

      var clientHtml =
        '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f4f3ef;font-family:Inter,system-ui,sans-serif">' +
        '<div style="max-width:580px;margin:0 auto;padding:24px 16px">' +

        '<div style="background:#0e0e0d;border-radius:12px 12px 0 0;padding:20px 24px;border-bottom:3px solid #5AB544">' +
        '<div style="font-size:16px;font-weight:700;color:#ffffff">Storm43 Creative Media Agency</div>' +
        '<div style="font-size:11px;color:rgba(255,255,255,0.35);font-family:monospace;margin-top:2px;letter-spacing:1px">BRIEF CONFIRMATION</div>' +
        '</div>' +

        '<div style="background:#ffffff;border:1px solid #e0dfd9;border-top:none;border-radius:0 0 12px 12px;padding:28px 24px">' +

        '<p style="font-size:15px;color:#18180f;margin-bottom:16px">Dear ' + (brief.clientName || 'Client') + ',</p>' +

        '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:20px">Thank you for submitting your project brief to <strong>Storm43 Creative Media Agency</strong>. We have received your brief and a member of our team will be in touch within <strong>1 business day</strong>.</p>' +

        '<div style="background:#f4f3ef;border-radius:8px;padding:16px 18px;margin-bottom:24px;border-left:3px solid #5AB544">' +
        '<div style="font-size:12px;color:#6b6b60;font-family:monospace;margin-bottom:8px;letter-spacing:0.5px">YOUR SUBMISSION DETAILS</div>' +
        '<div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:13px;color:#6b6b60">Reference</span><span style="font-size:13px;font-weight:700;color:#3B6D11;font-family:monospace">' + (brief.ref || '') + '</span></div>' +
        '<div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:13px;color:#6b6b60">Project</span><span style="font-size:13px;font-weight:600;color:#18180f">' + (brief.projectName || '') + '</span></div>' +
        '<div style="display:flex;justify-content:space-between"><span style="font-size:13px;color:#6b6b60">Design Type</span><span style="font-size:13px;color:#18180f">' + (brief.designType || '') + '</span></div>' +
        '</div>' +

        '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:20px">Our team will review your requirements and come back to you with a quotation and proposed timeline. If you have any urgent queries in the meantime, please reply to this email.</p>' +

        '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:24px">We look forward to working with you on this project.</p>' +

        '<p style="font-size:14px;color:#18180f;line-height:1.6">Kind regards,<br><strong>Storm43 Creative Media Agency</strong><br>' +
        '<span style="color:#6b6b60;font-size:12px">design@storm43.co.za</span></p>' +

        '</div>' +

        '<div style="text-align:center;padding:16px">' +
        '<div style="font-size:11px;color:#a8a89a;font-family:monospace">Storm43 Creative Media Agency &bull; All submissions are confidential</div>' +
        '</div>' +

        '</div></body></html>';

      var clientPlainText = 'Dear ' + (brief.clientName || 'Client') + ',\n\n' +
        'Thank you for submitting your project brief to Storm43 Creative Media Agency.\n\n' +
        'We have received your brief and a member of our team will be in touch within 1 business day.\n\n' +
        'Reference: ' + (brief.ref || '') + '\n' +
        'Project:   ' + (brief.projectName || '') + '\n\n' +
        'Kind regards,\nStorm43 Creative Media Agency\ndesign@storm43.co.za';

      MailApp.sendEmail({
        to: clientEmailAddr,
        subject: clientSubject,
        htmlBody: clientHtml,
        body: clientPlainText
      });

      Logger.log('Client acknowledgement sent successfully to: ' + clientEmailAddr);
    } else {
      Logger.log('Client email skipped — invalid or empty address: [' + clientEmailAddr + ']');
    }

  } catch (e) {
    Logger.log('Brief email error: ' + e.message);
  }
}


// ── Module-level HTML helpers for brief email ──
function briefRow(label, value) {
  if (!value) return '';
  return '<tr><td style="padding:8px 12px;font-size:12px;color:#6b6b60;font-family:monospace;width:160px;vertical-align:top;border-bottom:1px solid #f0efea">' + label + '</td>' +
    '<td style="padding:8px 12px;font-size:13px;color:#18180f;font-weight:500;border-bottom:1px solid #f0efea;vertical-align:top">' + value + '</td></tr>';
}

function briefSection(title, icon, rows) {
  return '<div style="margin-bottom:20px">' +
    '<div style="background:#0e0e0d;padding:10px 16px;border-radius:8px 8px 0 0;display:flex;align-items:center;gap:8px">' +
    '<span style="font-size:16px">' + icon + '</span>' +
    '<span style="color:#5AB544;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;font-family:monospace">' + title + '</span>' +
    '</div>' +
    '<table style="width:100%;border-collapse:collapse;background:#ffffff;border:1px solid #e0dfd9;border-top:none;border-radius:0 0 8px 8px;overflow:hidden">' +
    rows +
    '</table></div>';
}

// ── Send annotation report email to studio ──
function sendAnnotationReportEmail(params) {
  Logger.log('sendAnnotationReportEmail called');
  try {
    var jobRef    = params.jobRef    || '';
    var jobTitle  = params.jobTitle  || '';
    var decision  = params.decision  || '';
    var reviewer  = params.reviewer  || '';
    var client    = params.client    || '';
    var iteration = params.iteration || '1';
    var annCount  = params.annotations || '0';
    var reportHtml = params.reportHtml || '';
    var annSummary = params.annSummary || 'No annotations';

    var decisionLabels = { approved: 'APPROVED', rejected: 'REJECTED', revision: 'REVISIONS REQUESTED' };
    var label = decisionLabels[decision] || decision.toUpperCase();

    var subject = 'Annotation Report — ' + label + ' — ' + jobRef + ' — ' + jobTitle;

    // If we have the full HTML report, use it; otherwise build a simple one
    var htmlBody = reportHtml || (
      '<!DOCTYPE html><html><body style="font-family:Inter,sans-serif;padding:24px;background:#f4f3ef">' +
      '<div style="max-width:600px;margin:0 auto;background:#fff;border-radius:12px;padding:24px;border:1px solid #e0dfd9">' +
      '<h2 style="color:#18180f">Annotation Report — ' + jobRef + '</h2>' +
      '<p><strong>Decision:</strong> ' + label + '</p>' +
      '<p><strong>Project:</strong> ' + jobTitle + '</p>' +
      '<p><strong>Client:</strong> ' + client + '</p>' +
      '<p><strong>Reviewer:</strong> ' + reviewer + '</p>' +
      '<p><strong>Iteration:</strong> ' + iteration + '</p>' +
      '<p><strong>Annotations:</strong> ' + annCount + '</p>' +
      '<hr style="border:none;border-top:1px solid #e0dfd9;margin:16px 0">' +
      '<pre style="font-size:12px;color:#333;line-height:1.6">' + annSummary + '</pre>' +
      '</div></body></html>'
    );

    var plainBody = 'Annotation Report — ' + label + '\n' +
      'Project: ' + jobTitle + ' (' + jobRef + ')\n' +
      'Client: ' + client + '\n' +
      'Reviewer: ' + reviewer + '\n' +
      'Iteration: ' + iteration + '\n' +
      'Annotations: ' + annCount + '\n\n' +
      annSummary;

    MailApp.sendEmail({
      to: STUDIO_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody
    });

    Logger.log('Annotation report email sent to: ' + STUDIO_EMAIL);
    return { success: true };
  } catch(e) {
    Logger.log('Annotation report email error: ' + e.message);
    return { error: e.message };
  }
}

// Test the brief functions
function testBrief() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var testBrief = {
    ref: 'BRF-' + new Date().getFullYear() + '-TEST',
    submittedAt: new Date().toISOString(),
    clientName: 'Test Client',
    clientCompany: 'Test Company',
    clientEmail: STUDIO_EMAIL,
    clientPhone: '+27 11 000 0000',
    clientIndustry: 'Food & Beverage',
    projectName: 'Test Product Range',
    designType: 'Packaging',
    projectType: 'New',
    priority: 'High',
    dueDate: '2026-06-30',
    budget: 'R15,000 - R30,000',
    summary: 'This is a test brief submission to verify the system is working correctly.',
    audience: 'Health-conscious consumers aged 25-45',
    tone: 'Natural, Premium',
    colours: 'Earthy greens and warm terracotta',
    notes: 'Test submission - please ignore',
    attachedCount: 2,
    aiConversationSummary: 'USER: Help me write my brief\nASSISTANT: Sure, tell me about your project.'
  };

  var result = writeBriefData(ss, testBrief);
  Logger.log('writeBriefData result: ' + JSON.stringify(result));

  var readResult = readBriefs(ss);
  Logger.log('readBriefs count: ' + readResult.count);

  sendBriefNotificationEmail(testBrief);
  Logger.log('Test complete - check your email and the briefs/briefDetails sheets');
}

// ════════════════════════════════════════════════════════════
// ── EMAIL — EXISTING APPROVAL NOTIFICATIONS ──
// ════════════════════════════════════════════════════════════
function sendNotificationEmail(params) {
  Logger.log('sendNotificationEmail called. STUDIO_EMAIL=' + STUDIO_EMAIL);
  try {
    var decision    = params.decision    || 'unknown';
    var jobRef      = params.jobRef      || '';
    var jobTitle    = params.jobTitle    || '';
    var client      = params.client      || '';
    var reviewer    = params.reviewer    || '';
    var comment     = params.comment     || '';
    var annCount    = params.annotations || '0';
    var iteration   = params.iteration   || '1';
    var portalUrl   = params.portalUrl   || '';
    var clientEmail = params.clientEmail || '';

    var decisionLabels = { approved: 'APPROVED', rejected: 'REJECTED', revision: 'REVISIONS REQUESTED' };
    var decisionColors = { approved: '#3B6D11', rejected: '#991f1f', revision: '#c47d10' };
    var decisionBg     = { approved: '#eaf5e6',  rejected: '#fdf0f0',  revision: '#fdf0d6' };
    var decisionIcons  = { approved: '&#9989;',  rejected: '&#10060;', revision: '&#128221;' };

    var label     = decisionLabels[decision] || decision.toUpperCase();
    var labelColor = decisionColors[decision] || '#333';
    var labelBg    = decisionBg[decision]    || '#f4f4f4';
    var labelIcon  = decisionIcons[decision] || '&#128203;';
    var iter       = parseInt(iteration) > 1 ? ' (Iteration ' + iteration + ')' : '';
    var dateStr    = new Date().toLocaleDateString('en-ZA', {day:'2-digit', month:'long', year:'numeric', hour:'2-digit', minute:'2-digit'});
    var annNote    = parseInt(annCount) > 0 ? annCount + ' annotation' + (parseInt(annCount) !== 1 ? 's' : '') : 'None';

    // ════════════════════════════════════
    // STUDIO HTML EMAIL
    // ════════════════════════════════════
    var studioSubject = 'Design ' + label + iter + ' — ' + jobRef + ' — ' + jobTitle;

    var studioHtml =
      '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
      '<body style="margin:0;padding:0;background:#f4f3ef;font-family:Inter,system-ui,sans-serif">' +
      '<div style="max-width:620px;margin:0 auto;padding:24px 16px">' +

      // Header
      '<div style="background:#0e0e0d;border-radius:12px 12px 0 0;padding:20px 24px;display:flex;justify-content:space-between;align-items:center;border-bottom:3px solid #5AB544">' +
      '<div><div style="font-size:16px;font-weight:700;color:#fff">Storm43 MIS</div>' +
      '<div style="font-size:11px;color:rgba(255,255,255,0.35);font-family:monospace;letter-spacing:1px;margin-top:2px">DESIGN REVIEW NOTIFICATION</div></div>' +
      '<div style="text-align:right">' +
      '<div style="font-size:15px;font-weight:700;color:#5AB544;font-family:monospace">' + jobRef + '</div>' +
      '<div style="font-size:11px;color:rgba(255,255,255,0.3);font-family:monospace;margin-top:2px">' + dateStr + '</div>' +
      '</div></div>' +

      // Decision banner
      '<div style="background:' + labelBg + ';border:1px solid ' + labelColor + ';border-top:none;padding:16px 20px;margin-bottom:20px;border-radius:0 0 8px 8px;display:flex;align-items:center;gap:12px">' +
      '<span style="font-size:28px">' + labelIcon + '</span>' +
      '<div><div style="font-size:16px;font-weight:700;color:' + labelColor + '">' + label + iter + '</div>' +
      '<div style="font-size:12px;color:' + labelColor + ';opacity:0.8;margin-top:2px">Reviewed by ' + reviewer + '</div></div>' +
      '</div>' +

      // Project details card
      '<div style="background:#fff;border:1px solid #e0dfd9;border-radius:10px;overflow:hidden;margin-bottom:16px">' +
      '<div style="background:#f4f3ef;padding:10px 16px;border-bottom:1px solid #e0dfd9">' +
      '<span style="font-size:11px;font-weight:700;color:#6b6b60;font-family:monospace;letter-spacing:1px;text-transform:uppercase">Project Details</span>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse">' +
      '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;width:140px;border-bottom:1px solid #f0efea">Project</td><td style="padding:9px 16px;font-size:13px;font-weight:600;color:#18180f;border-bottom:1px solid #f0efea">' + jobTitle + '</td></tr>' +
      '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Reference</td><td style="padding:9px 16px;font-size:13px;font-weight:600;color:#3B6D11;font-family:monospace;border-bottom:1px solid #f0efea">' + jobRef + '</td></tr>' +
      '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Client</td><td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + client + '</td></tr>' +
      '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Reviewed By</td><td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + reviewer + '</td></tr>' +
      '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Iteration</td><td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + iteration + '</td></tr>' +
      '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace">Annotations</td><td style="padding:9px 16px;font-size:13px;color:#18180f">' + annNote + '</td></tr>' +
      '</table></div>' +

      // Comment block
      (comment ?
        '<div style="background:#fff;border:1px solid #e0dfd9;border-radius:10px;overflow:hidden;margin-bottom:16px">' +
        '<div style="background:#f4f3ef;padding:10px 16px;border-bottom:1px solid #e0dfd9">' +
        '<span style="font-size:11px;font-weight:700;color:#6b6b60;font-family:monospace;letter-spacing:1px;text-transform:uppercase">Client Comment</span>' +
        '</div>' +
        '<div style="padding:16px;font-size:13px;color:#18180f;line-height:1.7;border-left:3px solid ' + labelColor + ';margin:12px;background:' + labelBg + ';border-radius:6px">' +
        comment + '</div></div>' : '') +

      // Portal link
      (portalUrl ?
        '<div style="text-align:center;margin:20px 0">' +
        '<a href="' + portalUrl + '" style="display:inline-block;background:#1456a0;color:#fff;text-decoration:none;padding:11px 22px;border-radius:8px;font-size:13px;font-weight:600;margin-right:10px">View Design Portal &#8594;</a>' +
        '<a href="https://storm43design.github.io/Storm43-MIS/" style="display:inline-block;background:#5AB544;color:#fff;text-decoration:none;padding:11px 22px;border-radius:8px;font-size:13px;font-weight:600">Open MIS &#8594;</a>' +
        '</div>' :
        '<div style="text-align:center;margin:20px 0">' +
        '<a href="https://storm43design.github.io/Storm43-MIS/" style="display:inline-block;background:#5AB544;color:#fff;text-decoration:none;padding:11px 22px;border-radius:8px;font-size:13px;font-weight:600">Open Storm43 MIS &#8594;</a>' +
        '</div>') +

      // Footer
      '<div style="text-align:center;padding:14px">' +
      '<div style="font-size:11px;color:#a8a89a;font-family:monospace">Storm43 Creative Media Agency &bull; ' + jobRef + ' &bull; ' + dateStr + '</div>' +
      '</div>' +

      '</div></body></html>';

    var studioPlain = 'Storm43 MIS - Design Review Notification\n' +
      'Decision: ' + label + iter + '\n' +
      'Project: ' + jobTitle + ' (' + jobRef + ')\n' +
      'Client: ' + client + '\n' +
      'Reviewed by: ' + reviewer + '\n' +
      'Annotations: ' + annNote + '\n' +
      'Date: ' + dateStr + '\n' +
      (comment ? '\nClient comment:\n"' + comment + '"\n' : '') +
      '\nLog in to Storm43 MIS: https://storm43design.github.io/Storm43-MIS/';

    MailApp.sendEmail({
      to: STUDIO_EMAIL,
      subject: studioSubject,
      htmlBody: studioHtml,
      body: studioPlain
    });
    Logger.log('Studio HTML email sent to: ' + STUDIO_EMAIL);

    // ════════════════════════════════════
    // CLIENT HTML EMAIL
    // ════════════════════════════════════
    if (clientEmail && clientEmail.trim().length > 4 && clientEmail.indexOf('@') > 0 && clientEmail !== STUDIO_EMAIL) {
      var clientSubject = 'Design Review — ' + jobRef + ' — ' + jobTitle;

      var clientBodyContent = '';
      if (decision === 'approved') {
        clientBodyContent =
          '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:16px">Dear ' + reviewer + ',</p>' +
          '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:16px">Thank you for approving the design for <strong>' + jobTitle + '</strong>. Your approval has been recorded and our studio will now proceed with the next steps.</p>' +
          '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:16px">If you have any questions please do not hesitate to contact us.</p>';
      } else {
        var annText = parseInt(annCount) > 0
          ? ', including <strong>' + annCount + ' annotation' + (parseInt(annCount) !== 1 ? 's' : '') + '</strong>'
          : '';
        clientBodyContent =
          '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:16px">Dear ' + reviewer + ',</p>' +
          '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:16px">Thank you for reviewing the design for <strong>' + jobTitle + '</strong>. We have received your feedback' + annText + '.</p>' +
          (comment ?
            '<div style="background:#fdf0d6;border-left:3px solid #c47d10;border-radius:6px;padding:12px 16px;margin-bottom:16px;font-size:13px;color:#18180f;line-height:1.6">' +
            '<div style="font-size:10px;font-family:monospace;color:#c47d10;font-weight:700;letter-spacing:1px;margin-bottom:6px">YOUR COMMENT</div>' +
            comment + '</div>' : '') +
          '<p style="font-size:14px;color:#18180f;line-height:1.7;margin-bottom:16px">Our studio will review your comments and send you a revised design shortly.</p>' +
          (portalUrl ?
            '<div style="text-align:center;margin:20px 0">' +
            '<a href="' + portalUrl + '" style="display:inline-block;background:#1456a0;color:#fff;text-decoration:none;padding:11px 22px;border-radius:8px;font-size:13px;font-weight:600">View Design &#8594;</a>' +
            '</div>' : '');
      }

      var clientHtml =
        '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
        '<body style="margin:0;padding:0;background:#f4f3ef;font-family:Inter,system-ui,sans-serif">' +
        '<div style="max-width:580px;margin:0 auto;padding:24px 16px">' +

        '<div style="background:#0e0e0d;border-radius:12px 12px 0 0;padding:20px 24px;border-bottom:3px solid #5AB544">' +
        '<div style="font-size:16px;font-weight:700;color:#fff">Storm43 Creative Media Agency</div>' +
        '<div style="font-size:11px;color:rgba(255,255,255,0.35);font-family:monospace;letter-spacing:1px;margin-top:2px">DESIGN REVIEW</div>' +
        '</div>' +

        '<div style="background:#fff;border:1px solid #e0dfd9;border-top:none;border-radius:0 0 12px 12px;padding:28px 24px">' +

        // Decision badge
        '<div style="background:' + labelBg + ';border:1px solid ' + labelColor + ';border-radius:8px;padding:12px 16px;margin-bottom:20px;display:flex;align-items:center;gap:10px">' +
        '<span style="font-size:22px">' + labelIcon + '</span>' +
        '<div><div style="font-size:14px;font-weight:700;color:' + labelColor + '">' + label + '</div>' +
        '<div style="font-size:11px;color:' + labelColor + ';opacity:0.75;margin-top:1px">' + jobTitle + ' &bull; ' + jobRef + '</div></div>' +
        '</div>' +

        clientBodyContent +

        '<p style="font-size:14px;color:#18180f;line-height:1.6;margin-top:20px">Kind regards,<br>' +
        '<strong>Storm43 Creative Media Agency</strong><br>' +
        '<span style="color:#6b6b60;font-size:12px">design@storm43.co.za</span></p>' +

        '</div>' +

        '<div style="text-align:center;padding:14px">' +
        '<div style="font-size:11px;color:#a8a89a;font-family:monospace">Storm43 Creative Media Agency &bull; ' + jobRef + '</div>' +
        '</div>' +

        '</div></body></html>';

      var clientPlain = 'Dear ' + reviewer + ',\n\n' +
        (decision === 'approved' ?
          'Thank you for approving the design for ' + jobTitle + ' (' + jobRef + '). Our studio will proceed with the next steps.' :
          'Thank you for reviewing the design for ' + jobTitle + ' (' + jobRef + '). We have received your feedback and will send a revised design shortly.' +
          (comment ? '\n\nYour comment:\n"' + comment + '"' : '')) +
        '\n\nKind regards,\nStorm43 Creative Media Agency\ndesign@storm43.co.za';

      MailApp.sendEmail({
        to: clientEmail,
        subject: clientSubject,
        htmlBody: clientHtml,
        body: clientPlain
      });
      Logger.log('Client HTML email sent to: ' + clientEmail);
    }

    return { success: true, studioEmail: STUDIO_EMAIL, clientEmail: clientEmail };

  } catch (e) {
    Logger.log('Email error: ' + e.message);
    return { error: e.message };
  }
}

// ── TEST EMAIL ──
function testEmail() {
  Logger.log('Testing email. STUDIO_EMAIL = ' + STUDIO_EMAIL);
  var result = sendNotificationEmail({
    jobRef: 'STM-TEST-0001',
    jobTitle: 'Test Project',
    decision: 'revision',
    reviewer: 'Test Client',
    comment: 'Please adjust the font size on the label',
    annotations: '3',
    client: 'Test Client Co',
    clientEmail: STUDIO_EMAIL,
    iteration: '1',
    portalUrl: 'https://example.com/approval-portal.html?job=test123'
  });
  Logger.log('testEmail result: ' + JSON.stringify(result));
  Logger.log('Check your inbox at: ' + STUDIO_EMAIL);
}

// ── TEST SCRIPT ──

// ════════════════════════════════════════════════════════════
// FILE STORAGE — Google Drive (replaces chunk-based Sheets storage)
// ════════════════════════════════════════════════════════════
function saveToDrive(jobId, fileName, fileType, dataUrl) {
  try {
    if (!jobId || !dataUrl) return { error: 'Missing params' };

    var base64Data = dataUrl.split(',')[1] || dataUrl;
    var bytes = Utilities.base64Decode(base64Data);
    var mimeType = fileType === 'pdf' ? 'application/pdf' : 'image/jpeg';
    var uploadFileName = jobId + '_' + fileName;

    // Use Drive REST API via UrlFetchApp — avoids DriveApp scope issues
    var token = ScriptApp.getOAuthToken();

    // Step 1: Upload file content
    var uploadResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
      {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: {
          metadata: Utilities.newBlob(
            JSON.stringify({ name: uploadFileName, mimeType: mimeType }),
            'application/json'
          ),
          file: Utilities.newBlob(bytes, mimeType, uploadFileName)
        },
        muteHttpExceptions: true
      }
    );

    var uploadResult = JSON.parse(uploadResp.getContentText());
    if (!uploadResult.id) {
      Logger.log('Upload failed: ' + uploadResp.getContentText());
      return { error: 'Upload failed: ' + (uploadResult.error ? uploadResult.error.message : 'unknown') };
    }

    var fileId = uploadResult.id;

    // Step 2: Make file publicly readable
    UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId + '/permissions',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({ role: 'reader', type: 'anyone' }),
        muteHttpExceptions: true
      }
    );

    var viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
    Logger.log('File uploaded to Drive: ' + fileId);

    // Step 3: Store reference in Sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metaSheet = ss.getSheetByName('driveFiles');
    if (!metaSheet) {
      metaSheet = ss.insertSheet('driveFiles');
      metaSheet.getRange(1,1).setValue('meta');
    }
    var data = metaSheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === jobId) {
        metaSheet.getRange(i+1, 1, 1, 4).setValues([[jobId, fileId, viewUrl, new Date().toISOString()]]);
        found = true; break;
      }
    }
    if (!found) {
      metaSheet.appendRow([jobId, fileId, viewUrl, new Date().toISOString()]);
    }

    return { success: true, fileId: fileId, viewUrl: viewUrl };
  } catch(e) {
    Logger.log('saveToDrive error: ' + e.message);
    return { error: e.message };
  }
}

function getFileFromDrive(jobId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metaSheet = ss.getSheetByName('driveFiles');
    if (!metaSheet) return { hasFile: false, error: 'No files uploaded yet' };

    var data = metaSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === jobId) {
        var fileId = data[i][1];
        // Use REST API to get file metadata — no DriveApp scope needed
        var token = ScriptApp.getOAuthToken();
        var metaResp = UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=name,mimeType',
          { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
        );
        var meta = JSON.parse(metaResp.getContentText());
        var fileType = (meta.mimeType || '').indexOf('pdf') >= 0 ? 'pdf' : 'image';
        return {
          hasFile: true, jobId: jobId,
          fileId: fileId,
          fileName: meta.name || 'Design',
          fileType: fileType
        };
      }
    }
    return { hasFile: false, error: 'No file found for this job' };
  } catch(e) {
    Logger.log('getFileFromDrive error: ' + e.message);
    return { hasFile: false, error: e.message };
  }
}

function testDriveAccess() {
  try {
    var token = ScriptApp.getOAuthToken();
    // Test upload via REST API
    var testBytes = Utilities.base64Decode(Utilities.base64Encode('test'));
    var resp = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
      {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: {
          metadata: Utilities.newBlob(JSON.stringify({name:'_storm43_test.txt'}), 'application/json'),
          file: Utilities.newBlob('test content', 'text/plain', '_storm43_test.txt')
        },
        muteHttpExceptions: true
      }
    );
    var result = JSON.parse(resp.getContentText());
    if (result.id) {
      Logger.log('Drive REST API write access confirmed. File ID: ' + result.id);
      // Delete test file
      UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + result.id,
        { method: 'DELETE', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
      Logger.log('Test file cleaned up');
    } else {
      Logger.log('Drive REST API error: ' + resp.getContentText());
    }
    return result.id ? 'Drive access OK' : 'Failed: ' + resp.getContentText();
  } catch(e) {
    Logger.log('testDriveAccess error: ' + e.message);
    return e.message;
  }
}


// ════════════════════════════════════════════════════════════
// CRM — SAVE LEAD
// ════════════════════════════════════════════════════════════
function saveLead(ss, lead) {
  try {
    if (!lead || !lead.name || !lead.email) return { error: 'Missing required fields' };

    // Get or create leads sheet
    var sheet = ss.getSheetByName('leads');
    if (!sheet) {
      sheet = ss.insertSheet('leads');
      sheet.getRange(1,1,1,9).setValues([[
        'id','name','company','email','phone','source','needs','stage','createdAt'
      ]]);
    }

    // Append new lead
    sheet.appendRow([
      lead.id || ('lead_' + Date.now()),
      lead.name, lead.company || '', lead.email,
      lead.phone || '', lead.source || 'Other',
      lead.industry || '', lead.challenge || '',
      lead.needs || '', lead.stage || 'New',
      lead.createdAt || new Date().toISOString()
    ]);

    // Send notification email to studio
    try {
      MailApp.sendEmail({
        to: STUDIO_EMAIL,
        subject: '🎯 New Lead: ' + lead.name + ' (' + (lead.company||'Individual') + ')',
        htmlBody: '<div style="font-family:sans-serif;max-width:520px">'
          + '<div style="background:#1a1a1a;padding:20px;border-radius:8px 8px 0 0">'
          + '<h2 style="color:#4CAF50;margin:0">storm43 — New Lead</h2></div>'
          + '<div style="background:#f9f9f9;padding:20px;border:1px solid #e0e0e0;border-radius:0 0 8px 8px">'
          + '<table style="width:100%;border-collapse:collapse">'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px;width:120px">Name</td><td style="padding:6px 0;font-weight:700">' + lead.name + '</td></tr>'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px">Company</td><td style="padding:6px 0">' + (lead.company||'—') + '</td></tr>'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px">Email</td><td style="padding:6px 0"><a href="mailto:' + lead.email + '">' + lead.email + '</a></td></tr>'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px">Phone</td><td style="padding:6px 0">' + (lead.phone||'—') + '</td></tr>'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px">Source</td><td style="padding:6px 0">' + lead.source + '</td></tr>'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px">Industry</td><td style="padding:6px 0">' + (lead.industry||'—') + '</td></tr>'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px;vertical-align:top">Current Challenge</td><td style="padding:6px 0">' + (lead.challenge||'—') + '</td></tr>'
          + '<tr><td style="padding:6px 0;color:#666;font-size:13px;vertical-align:top">What they need</td><td style="padding:6px 0">' + lead.needs + '</td></tr>'
          + '</table>'
          + '<div style="margin-top:16px;padding-top:16px;border-top:1px solid #e0e0e0">'
          + '<a href="https://storm43design.github.io/Storm43-MIS/" style="background:#4CAF50;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:700;font-size:13px">Open MIS → View Lead</a>'
          + '</div></div></div>'
      });
    } catch(emailErr) {
      Logger.log('Lead email error: ' + emailErr.message);
    }

    return { success: true, leadId: lead.id };
  } catch(e) {
    Logger.log('saveLead error: ' + e.message);
    return { error: e.message };
  }
}

function testRead() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var result = readAllData(ss);
    Logger.log('Read OK. Jobs: ' + (result.jobs||[]).length);
    Logger.log('Users: ' + (result.users||[]).length);
    Logger.log('JobCounter: ' + result.jobCounter);
    return 'OK';
  } catch(e) {
    Logger.log('testRead ERROR: ' + e.message);
    return e.message;
  }
}

function testGetFile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // List all sheets to see what file sheets exist
  var sheets = ss.getSheets().map(function(sh){ return sh.getName(); });
  Logger.log('All sheets: ' + JSON.stringify(sheets));
  
  // Find file sheets (they start with 'f_')
  var fileSheets = sheets.filter(function(n){ return n.indexOf('f_') === 0; });
  Logger.log('File sheets: ' + JSON.stringify(fileSheets));
  
  // Try to get the first one
  if (fileSheets.length > 0) {
    var jobId = fileSheets[0].replace('f_', '');
    Logger.log('Testing getFile for jobId: ' + jobId);
    var result = getFile(ss, jobId);
    Logger.log('hasFile: ' + result.hasFile);
    Logger.log('fileType: ' + result.fileType);
    Logger.log('dataUrl length: ' + (result.dataUrl ? result.dataUrl.length : 0));
    Logger.log('error: ' + (result.error || 'none'));
  } else {
    Logger.log('No file sheets found — files have not been uploaded to Sheets yet');
  }
}

function testScript() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Sheets access OK: ' + ss.getName());

  var r1 = saveFileChunk(ss, 'test_job_999', 0, 2, 'CHUNK_ONE', 'pdf', 'test.pdf');
  Logger.log('Chunk 0: ' + JSON.stringify(r1));
  var r2 = saveFileChunk(ss, 'test_job_999', 1, 2, 'CHUNK_TWO', 'pdf', 'test.pdf');
  Logger.log('Chunk 1: ' + JSON.stringify(r2));

  var rf = getFile(ss, 'test_job_999');
  Logger.log('Retrieved: ' + rf.dataUrl);
  Logger.log('Match: ' + (rf.dataUrl === 'CHUNK_ONECHUNK_TWO'));

  deleteFile(ss, 'test_job_999');
  Logger.log('All tests complete');
}

// ── AUTHORISE EMAIL (run this ONCE to grant Gmail permission) ──
function authoriseEmail() {
  MailApp.sendEmail({
    to: STUDIO_EMAIL,
    subject: 'Storm43 MIS — Email Authorised',
    body: 'Email permission granted successfully. You can now receive approval notifications.'
  });
  Logger.log('Email authorised and test sent to: ' + STUDIO_EMAIL);
}

// ── Send assignment notification email to designer ──
function sendDesignerAssignmentEmail(params) {
  Logger.log('sendDesignerAssignmentEmail called for: ' + (params.designerEmail || 'NO EMAIL'));
  try {
    if (!params.designerEmail || params.designerEmail.indexOf('@') < 0) {
      return { error: 'Invalid designer email: ' + params.designerEmail };
    }

    var designerName = params.designerName || 'Designer';
    var jobRef       = params.jobRef      || '';
    var jobTitle     = params.jobTitle    || '';
    var jobClient    = params.jobClient   || '—';
    var jobDesignType= params.jobDesignType || '—';
    var jobProjectType=params.jobProjectType|| '—';
    var jobStage     = params.jobStage    || '—';
    var jobPriority  = params.jobPriority || '—';
    var jobDueDate   = params.jobDueDate  || 'Not specified';
    var jobSummary   = params.jobSummary  || '';
    var jobBrand     = params.jobBrand    || '';
    var dateStr      = new Date().toLocaleDateString('en-ZA', {day:'2-digit', month:'long', year:'numeric'});
    var subject      = 'Project Assigned — ' + jobRef + ' — ' + jobTitle;

    // Build branded HTML email in Apps Script (avoids JSON serialisation issues)
    var htmlBody =
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<style>*{box-sizing:border-box}body{margin:0;padding:0;background:#f4f3ef;font-family:Inter,system-ui,sans-serif}</style>' +
      '</head><body>' +
      '<div style="max-width:620px;margin:0 auto;padding:24px 16px">' +

      // Header
      '<div style="background:#0e0e0d;border-radius:12px 12px 0 0;padding:20px 24px;display:flex;justify-content:space-between;align-items:center;border-bottom:3px solid #5AB544">' +
      '<div>' +
        '<div style="font-size:16px;font-weight:700;color:#fff;letter-spacing:-0.3px">Storm43 MIS</div>' +
        '<div style="font-size:11px;color:rgba(255,255,255,0.35);font-family:monospace;letter-spacing:1px;margin-top:2px">PROJECT ASSIGNED</div>' +
      '</div>' +
      '<div style="text-align:right">' +
        '<div style="font-size:15px;font-weight:700;color:#5AB544;font-family:monospace">' + jobRef + '</div>' +
        '<div style="font-size:11px;color:rgba(255,255,255,0.3);font-family:monospace;margin-top:2px">' + dateStr + '</div>' +
      '</div>' +
      '</div>' +

      // Green assignment banner
      '<div style="background:#eaf5e6;border:1px solid #c8e8c0;border-top:none;padding:16px 20px;border-radius:0 0 8px 8px;margin-bottom:20px;display:flex;align-items:center;gap:14px">' +
      '<span style="font-size:32px">🎨</span>' +
      '<div>' +
        '<div style="font-size:16px;font-weight:700;color:#3B6D11">You have been assigned a new project</div>' +
        '<div style="font-size:12px;color:#3B6D11;opacity:0.8;margin-top:3px">Hi ' + designerName + ', please review the details below and get started</div>' +
      '</div>' +
      '</div>' +

      // Project details card
      '<div style="background:#fff;border:1px solid #e0dfd9;border-radius:10px;overflow:hidden;margin-bottom:16px">' +
        '<div style="background:#f4f3ef;padding:10px 16px;border-bottom:1px solid #e0dfd9">' +
          '<span style="font-size:11px;font-weight:700;color:#6b6b60;font-family:monospace;letter-spacing:1px;text-transform:uppercase">Project Details</span>' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse">' +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;width:140px;border-bottom:1px solid #f0efea">Project</td>' +
               '<td style="padding:9px 16px;font-size:13px;font-weight:700;color:#18180f;border-bottom:1px solid #f0efea">' + jobTitle + '</td></tr>' +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Reference</td>' +
               '<td style="padding:9px 16px;font-size:13px;font-weight:700;color:#3B6D11;font-family:monospace;border-bottom:1px solid #f0efea">' + jobRef + '</td></tr>' +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Client</td>' +
               '<td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + jobClient + '</td></tr>' +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Design Type</td>' +
               '<td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + jobDesignType + '</td></tr>' +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Project Type</td>' +
               '<td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + jobProjectType + '</td></tr>' +
          (jobBrand ? '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Brand</td>' +
               '<td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + jobBrand + '</td></tr>' : '') +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Stage</td>' +
               '<td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + jobStage + '</td></tr>' +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace;border-bottom:1px solid #f0efea">Priority</td>' +
               '<td style="padding:9px 16px;font-size:13px;color:#18180f;border-bottom:1px solid #f0efea">' + jobPriority + '</td></tr>' +
          '<tr><td style="padding:9px 16px;font-size:12px;color:#6b6b60;font-family:monospace">Due Date</td>' +
               '<td style="padding:9px 16px;font-size:13px;color:#18180f">' + jobDueDate + '</td></tr>' +
        '</table>' +
      '</div>' +

      // Brief summary
      (jobSummary ?
        '<div style="background:#fff;border:1px solid #e0dfd9;border-radius:10px;overflow:hidden;margin-bottom:16px">' +
          '<div style="background:#f4f3ef;padding:10px 16px;border-bottom:1px solid #e0dfd9">' +
            '<span style="font-size:11px;font-weight:700;color:#6b6b60;font-family:monospace;letter-spacing:1px;text-transform:uppercase">Project Brief</span>' +
          '</div>' +
          '<div style="padding:14px 16px;font-size:13px;color:#18180f;line-height:1.7;border-left:3px solid #5AB544;margin:12px;background:#f4f3ef;border-radius:6px">' + jobSummary + '</div>' +
        '</div>' : '') +

      // CTA button
      '<div style="text-align:center;margin:24px 0">' +
        '<a href="https://storm43design.github.io/Storm43-MIS/" ' +
           'style="display:inline-block;background:#5AB544;color:#ffffff;text-decoration:none;' +
                  'padding:13px 32px;border-radius:8px;font-size:14px;font-weight:700;letter-spacing:0.3px">' +
          'Open Storm43 MIS &#8594;' +
        '</a>' +
      '</div>' +

      // Footer
      '<div style="text-align:center;padding:14px;border-top:1px solid #e0dfd9">' +
        '<div style="font-size:11px;color:#a8a89a;font-family:monospace">' +
          'Storm43 Creative Media Agency &bull; ' + jobRef + ' &bull; ' + dateStr +
        '</div>' +
      '</div>' +

      '</div></body></html>';

    var nl = '\n';
    var plainBody =
      'Hi ' + designerName + ',' + nl + nl +
      'You have been assigned to a new project.' + nl + nl +
      'Project:      ' + jobTitle + nl +
      'Reference:    ' + jobRef + nl +
      'Client:       ' + jobClient + nl +
      'Design Type:  ' + jobDesignType + nl +
      'Project Type: ' + jobProjectType + nl +
      'Stage:        ' + jobStage + nl +
      'Priority:     ' + jobPriority + nl +
      'Due Date:     ' + jobDueDate + nl +
      (jobSummary ? nl + 'Brief:' + nl + jobSummary + nl : '') +
      nl + 'Log in to Storm43 MIS:' + nl +
      'https://storm43design.github.io/Storm43-MIS/' + nl + nl +
      'Kind regards,' + nl + 'Storm43 Creative Media Agency';

    MailApp.sendEmail({
      to: params.designerEmail,
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody
    });
    Logger.log('Assignment email sent to designer: ' + params.designerEmail);

    // Studio confirmation (plain text only)
    MailApp.sendEmail({
      to: STUDIO_EMAIL,
      subject: 'Designer Notified — ' + jobRef + ' assigned to ' + designerName,
      body: designerName + ' (' + params.designerEmail + ') has been notified for project ' + jobRef + ' - ' + jobTitle
    });

    return { success: true, sentTo: params.designerEmail };
  } catch (e) {
    Logger.log('Assignment email error: ' + e.message);
    return { error: e.message };
  }
}
