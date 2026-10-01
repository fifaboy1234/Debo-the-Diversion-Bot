const CONFIG = {
  recipients: [
    'shariff.bakar@tbiam.aero',
  ],

  timezone: 'America/Los_Angeles',

  kontAirport: 'KONT',

  sourceAirports: [
    'KSAN',
    'KLAS',
    'KPSP',
    'KPHX',
    'KBUR'
  ],

  sheetName: 'KONT Diversion Watch',
  logSheetName: 'Daily_Log',

  metarUrl:
    'https://aviationweather.gov/api/data/metar?format=json&hours=3&ids=',

  tafUrl:
    'https://aviationweather.gov/api/data/taf?format=json&ids=',

  nasStatusUrl:
    'https://nasstatus.faa.gov/list',

  awcLink:
    'https://aviationweather.gov/data/metar/?ids=KONT%2CKSAN%2CKLAS%2CKPSP%2CKPHX%2CKBUR&taf=1',

  nasLink:
    'https://nasstatus.faa.gov/',

  delayLink:
    'https://www.fly.faa.gov/fly/flyfaa/nemap',

  notamLink:
    'https://notams.aim.faa.gov/notamSearch/'
};

function setup() {
  const spreadsheet = getOrCreateSpreadsheet_();

  createSheetIfNeeded_(
    spreadsheet,
    CONFIG.logSheetName,
    [
      'Issued local time',
      'Issued Zulu time',
      'Overall rating',
      'NAS status',
      'Airport',
      'Airport rating',
      'Drivers',
      'METAR',
      'TAF'
    ]
  );
}

function sendDailyDiversionWatch() {
  const report = buildReport_();

  saveReportToLog_(report);

  MailApp.sendEmail({
    to: CONFIG.recipients.join(','),
    subject:
      'KONT Diversion Watch — ' +
      report.overallRating +
      ' — ' +
      report.issuedLocal,
    htmlBody: report.html,
    body: report.text
  });
}

function testEmailNow() {
  sendDailyDiversionWatch();
}

function createDaily6amTrigger() {
  removeExistingTrigger_('sendDailyDiversionWatch');

  ScriptApp.newTrigger('sendDailyDiversionWatch')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone(CONFIG.timezone)
    .create();
}

function buildReport_() {
  const now = new Date();

  const issuedLocal = Utilities.formatDate(
    now,
    CONFIG.timezone,
    'EEE MMM d, yyyy h:mm a z'
  );

  const issuedZulu = Utilities.formatDate(
    now,
    'UTC',
    "yyyy-MM-dd HH:mm 'Z'"
  );

  const allAirportIds = [
    CONFIG.kontAirport
  ].concat(CONFIG.sourceAirports);

  const metars = getJson_(
    CONFIG.metarUrl + allAirportIds.join(',')
  );

  const tafs = getJson_(
    CONFIG.tafUrl + allAirportIds.join(',')
  );

  const metarsByAirport = mapAirportReports_(metars);
  const tafsByAirport = mapAirportReports_(tafs);

  const kont = evaluateAirport_(
    CONFIG.kontAirport,
    metarsByAirport[CONFIG.kontAirport] || null,
    tafsByAirport[CONFIG.kontAirport] || null
  );

  const sourceAirports = CONFIG.sourceAirports.map(function(airportId) {
    return evaluateAirport_(
      airportId,
      metarsByAirport[airportId] || null,
      tafsByAirport[airportId] || null
    );
  });

  const nasStatus = getNasStatus_(allAirportIds);

  let highestScore = 0;

  sourceAirports.forEach(function(airport) {
    if (airport.internalScore > highestScore) {
      highestScore = airport.internalScore;
    }
  });

  if (nasStatus.hasAffectedAirports) {
    highestScore = Math.max(
      highestScore,
      nasStatus.highestSeverity
    );
  }

  let overallRating = 'LOW';

  if (highestScore >= 6) {
    overallRating = 'ELEVATED';
  } else if (highestScore >= 3) {
    overallRating = 'WATCH';
  }

  const report = {
    issuedLocal: issuedLocal,
    issuedZulu: issuedZulu,
    overallRating: overallRating,
    kont: kont,
    sourceAirports: sourceAirports,
    allAirports: [kont].concat(sourceAirports),
    nasStatus: nasStatus
  };

  report.html = createHtmlEmail_(report);
  report.text = createPlainTextEmail_(report);

  return report;
}

function evaluateAirport_(airportId, metar, taf) {
  const metarText = getRawReportText_(metar);
  const tafText = getRawReportText_(taf);

  const combinedText = (
    metarText +
    ' ' +
    tafText
  ).toUpperCase();

  let internalScore = 0;
  const drivers = [];

  const hasThunderstorm = /\b(TSRA|VCTS|TS\b|CB\b)/.test(
    combinedText
  );

  const hasWindShiftOrShear = /\b(WS\d|WSHFT|LLWS)\b/.test(
    combinedText
  );

  const hasGust25OrGreater = /G2[5-9]KT|G[3-9]\dKT/.test(
    combinedText
  );

  const hasRecentPeakWind25OrGreater =
    /PK WND \d{3}(2[5-9]|[3-9]\d)\//.test(
      combinedText
    );

  const hasRecentRainEnded = /\bRAE\d{2}\b/.test(
    combinedText
  );

  const hasPrecipitation = /\b(RA|SN|FZRA|DZ)\b/.test(
    combinedText
  );

  /*
   * Ceiling detection:
   *
   * BKN006 = broken ceiling at 600 ft
   * OVC004 = overcast ceiling at 400 ft
   * VV002 = vertical visibility of 200 ft
   *
   * FEW006 and SCT006 do not count as ceilings.
   */
  const hasCeilingBelow1000 = /\b(BKN|OVC|VV)0[0-9]{2}\b/.test(
    combinedText
  );

  const hasCeiling500OrLower = /\b(BKN|OVC|VV)00[0-5]\b/.test(
    combinedText
  );

  /*
   * Visibility below 3 SM:
   *
   * M1/4SM
   * 1/4SM
   * 1/2SM
   * 3/4SM
   * 1SM
   * 2SM
   */
  const hasVisibilityBelow3 =
    /\b(M1\/4SM|[1-2]\/[248]SM|3\/4SM|1SM|2SM)\b/.test(
      combinedText
    );

  const hasVisibilityBelow1 =
    /\b(M1\/4SM|[1-2]\/[248]SM|3\/4SM)\b/.test(
      combinedText
    );

  if (hasThunderstorm) {
    internalScore += 6;
    drivers.push('thunderstorm or convective signal');
  }

  if (hasWindShiftOrShear) {
    internalScore += 2;
    drivers.push('wind shear or wind shift');
  }

  if (hasGust25OrGreater) {
    internalScore += 3;
    drivers.push('gusts 25 kt or greater');
  }

  if (hasRecentPeakWind25OrGreater) {
    internalScore += 3;
    drivers.push('recent peak wind 25 kt or greater');
  }

  /*
   * A ceiling below 1,000 ft always reaches
   * the Watch threshold by itself.
   */
  if (hasCeilingBelow1000) {
    internalScore += 3;
    drivers.push('ceiling below 1,000 ft');
  }

  if (hasCeiling500OrLower) {
    internalScore += 3;
    drivers.push('ceiling at or below 500 ft');
  }

  if (hasVisibilityBelow3) {
    internalScore += 2;
    drivers.push('visibility below 3 SM');
  }

  if (hasVisibilityBelow1) {
    internalScore += 4;
    drivers.push('visibility below 1 SM');
  }

  if (hasPrecipitation) {
    internalScore += 1;
    drivers.push('precipitation');
  }

  if (hasRecentRainEnded) {
    internalScore += 1;
    drivers.push('recent rain ended');
  }

  let rating = 'LOW';

  if (internalScore >= 6) {
    rating = 'ELEVATED';
  } else if (internalScore >= 3) {
    rating = 'WATCH';
  }

  return {
    id: airportId,
    metar: metarText,
    taf: tafText,
    drivers: drivers.length
      ? drivers
      : ['no automated threshold met'],
    internalScore: internalScore,
    rating: rating
  };
}

function getNasStatus_(airportIds) {
  const fallback = {
    hasAffectedAirports: false,
    affectedAirports: [],
    highestSeverity: 0,
    displayText: 'NONE',
    retrievalMessage: ''
  };

  let response;

  try {
    response = UrlFetchApp.fetch(CONFIG.nasStatusUrl, {
      muteHttpExceptions: true,
      headers: {
        Accept: 'text/html',
        'User-Agent': 'Mozilla/5.0'
      }
    });
  } catch (error) {
    fallback.retrievalMessage =
      'FAA NAS Status could not be retrieved automatically. Open the FAA NAS Status link to verify.';
    return fallback;
  }

  if (response.getResponseCode() !== 200) {
    fallback.retrievalMessage =
      'FAA NAS Status could not be retrieved automatically. Open the FAA NAS Status link to verify.';
    return fallback;
  }

  const html = response.getContentText();
  const plainText = htmlToPlainText_(html);

  const lines = plainText
    .split(/\n+/)
    .map(function(line) {
      return line.replace(/\s+/g, ' ').trim();
    })
    .filter(function(line) {
      return line.length > 0;
    });

  const aliases = {
    KONT: ['KONT', 'ONT'],
    KSAN: ['KSAN', 'SAN'],
    KLAS: ['KLAS', 'LAS'],
    KPSP: ['KPSP', 'PSP'],
    KPHX: ['KPHX', 'PHX'],
    KBUR: ['KBUR', 'BUR']
  };

  const affectedAirports = [];

  airportIds.forEach(function(airportId) {
    const airportAliases = aliases[airportId] || [airportId];
    const matchingEvents = [];

    lines.forEach(function(line) {
      const uppercaseLine = line.toUpperCase();

      const airportMentioned = airportAliases.some(function(alias) {
        const pattern = new RegExp(
          '(^|[^A-Z0-9])' +
          escapeRegex_(alias) +
          '([^A-Z0-9]|$)'
        );

        return pattern.test(uppercaseLine);
      });

      const flowEventMentioned =
        /\b(GROUND STOP|GROUND DELAY PROGRAM|GROUND DELAY|GDP|AIRSPACE FLOW PROGRAM|AFP|DELAY|ARRIVAL RATE|CLOSURE|TRAFFIC MANAGEMENT)\b/.test(
          uppercaseLine
        );

      if (airportMentioned && flowEventMentioned) {
        matchingEvents.push(line);
      }
    });

    const uniqueEvents = unique_(matchingEvents).slice(0, 3);

    if (uniqueEvents.length > 0) {
      affectedAirports.push({
        id: airportId,
        events: uniqueEvents,
        severity: getNasSeverity_(uniqueEvents.join(' '))
      });
    }
  });

  if (affectedAirports.length === 0) {
    return fallback;
  }

  let highestSeverity = 0;

  affectedAirports.forEach(function(airport) {
    if (airport.severity > highestSeverity) {
      highestSeverity = airport.severity;
    }
  });

  return {
    hasAffectedAirports: true,
    affectedAirports: affectedAirports,
    highestSeverity: highestSeverity,
    displayText: affectedAirports
      .map(function(airport) {
        return airport.id + ': ' + airport.events.join(' | ');
      })
      .join(' — '),
    retrievalMessage: ''
  };
}

function getNasSeverity_(text) {
  const value = String(text).toUpperCase();

  if (value.indexOf('GROUND STOP') !== -1) {
    return 6;
  }

  if (
    value.indexOf('GROUND DELAY PROGRAM') !== -1 ||
    value.indexOf('GDP') !== -1
  ) {
    return 5;
  }

  if (
    value.indexOf('GROUND DELAY') !== -1 ||
    value.indexOf('AIRSPACE FLOW PROGRAM') !== -1 ||
    value.indexOf('AFP') !== -1 ||
    value.indexOf('CLOSURE') !== -1
  ) {
    return 4;
  }

  return 3;
}

function getJson_(url) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      Accept: 'application/json'
    }
  });

  if (response.getResponseCode() !== 200) {
    return [];
  }

  try {
    return JSON.parse(response.getContentText());
  } catch (error) {
    return [];
  }
}

function mapAirportReports_(reports) {
  const output = {};

  (reports || []).forEach(function(report) {
    const airportId = String(
      report.icaoId ||
      report.stationId ||
      report.id ||
      ''
    ).toUpperCase();

    if (airportId) {
      output[airportId] = report;
    }
  });

  return output;
}

function getRawReportText_(report) {
  if (!report) {
    return 'Unavailable from automated feed';
  }

  return String(
    report.rawOb ||
    report.rawTAF ||
    report.rawText ||
    report.text ||
    report.report ||
    JSON.stringify(report)
  );
}

function htmlToPlainText_(html) {
  return String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<\/(p|div|tr|li|h1|h2|h3|br)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>');
}

function getOrCreateSpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  const matchingFiles = DriveApp.getFilesByName(
    CONFIG.sheetName
  );

  if (matchingFiles.hasNext()) {
    return SpreadsheetApp.open(matchingFiles.next());
  }

  return SpreadsheetApp.create(CONFIG.sheetName);
}

function createSheetIfNeeded_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#d9eaf7');

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

function createHtmlEmail_(report) {
  let sourceRows = '';

  report.sourceAirports.forEach(function(airport) {
    const nasAirport = findNasAirport_(
      report.nasStatus,
      airport.id
    );

    let nasText = '';

    if (nasAirport) {
      nasText =
        '<br>' +
        '<span style="font-size:12px;color:#8a4100;">' +
        '<b>FAA NAS:</b> ' +
        escapeHtml_(nasAirport.events.join(' | ')) +
        '</span>';
    }

    sourceRows +=
      '<tr>' +
      '<td style="padding:12px 10px;border-bottom:1px solid #dbe4ea;vertical-align:top;">' +
      '<b>' + airport.id + '</b>' +
      '</td>' +
      '<td style="padding:12px 10px;border-bottom:1px solid #dbe4ea;vertical-align:top;">' +
      createStatusBox_(airport.rating) +
      '</td>' +
      '<td style="padding:12px 10px;border-bottom:1px solid #dbe4ea;vertical-align:top;">' +
      escapeHtml_(airport.drivers.join('; ')) +
      nasText +
      '</td>' +
      '</tr>';
  });

  let airportDetails = '';

  report.allAirports.forEach(function(airport) {
    const nasAirport = findNasAirport_(
      report.nasStatus,
      airport.id
    );

    let nasDetail = '';

    if (nasAirport) {
      nasDetail =
        '<br><b>FAA NAS:</b> ' +
        escapeHtml_(nasAirport.events.join(' | '));
    }

    airportDetails +=
      '<div style="margin:0 0 20px 0;">' +
      '<div style="margin:0 0 8px 0;font-size:17px;font-weight:bold;">' +
      airport.id +
      '&nbsp;&nbsp;' +
      createStatusBox_(airport.rating) +
      '</div>' +
      '<div style="line-height:1.6;">' +
      '<b>Drivers:</b> ' +
      escapeHtml_(airport.drivers.join('; ')) +
      nasDetail +
      '<br><b>METAR:</b> ' +
      escapeHtml_(airport.metar) +
      '<br><b>TAF:</b> ' +
      escapeHtml_(airport.taf) +
      '</div>' +
      '</div>';
  });

  return (
    '<div style="font-family:Arial,sans-serif;max-width:950px;color:#152535;padding:8px;">' +

    '<h1 style="margin:0 0 16px 0;font-size:26px;">KONT Diversion Watch</h1>' +

    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" ' +
    'style="border-collapse:collapse;margin:0 0 20px 0;">' +

    '<tr>' +
    '<td style="padding:0 14px 10px 0;font-weight:bold;white-space:nowrap;">Issued:</td>' +
    '<td style="padding:0 0 10px 0;">' +
    report.issuedLocal +
    ' (' +
    report.issuedZulu +
    ')' +
    '</td>' +
    '</tr>' +

    '<tr>' +
    '<td style="padding:0 14px 10px 0;font-weight:bold;white-space:nowrap;">Overall rating:</td>' +
    '<td style="padding:0 0 10px 0;">' +
    createStatusBox_(report.overallRating) +
    '</td>' +
    '</tr>' +

    '<tr>' +
    '<td style="padding:0 14px 0 0;font-weight:bold;white-space:nowrap;">NAS Status:</td>' +
    '<td style="padding:0;">' +
    createNasStatusBox_(report.nasStatus) +
    '</td>' +
    '</tr>' +

    '</table>' +

    '<hr style="border:none;border-top:1px solid #dbe4ea;margin:0 0 22px 0;">' +

    '<h2 style="margin:0 0 12px 0;font-size:20px;">Source-airport risk</h2>' +

    '<table cellpadding="0" cellspacing="0" border="0" ' +
    'style="border-collapse:collapse;width:100%;margin:0 0 24px 0;">' +

    '<tr style="background:#e8eef3;">' +
    '<th align="left" style="padding:11px 10px;">Airport</th>' +
    '<th align="left" style="padding:11px 10px;">Rating</th>' +
    '<th align="left" style="padding:11px 10px;">Weather / operational drivers</th>' +
    '</tr>' +

    sourceRows +

    '</table>' +

    '<h2 style="margin:0 0 10px 0;font-size:20px;">Operational links</h2>' +

    '<ul style="margin:0 0 24px 0;padding-left:20px;line-height:1.7;">' +
    '<li><a href="' + CONFIG.awcLink + '">' +
    'Aviation Weather Center: six-airport METARs and TAFs' +
    '</a></li>' +
    '<li><a href="' + CONFIG.nasLink + '">' +
    'FAA NAS Status' +
    '</a></li>' +
    '<li><a href="' + CONFIG.delayLink + '">' +
    'FAA Delay Information' +
    '</a></li>' +
    '<li><a href="' + CONFIG.notamLink + '">' +
    'FAA NOTAM Search' +
    '</a></li>' +
    '</ul>' +

    '<h2 style="margin:0 0 12px 0;font-size:20px;">Airport detail</h2>' +

    airportDetails +

    '<hr style="border:none;border-top:1px solid #dbe4ea;margin:4px 0 16px 0;">' +

    '<p style="margin:0;font-size:12px;line-height:1.5;color:#526473;">' +
    '<i>' +
    'Decision support only—not a Part 121 dispatch release. ' +
    'Verify FAA flow programs, NOTAMs, radar/lightning, ADS-B behavior, ' +
    'and KONT operational capacity before action.' +
    '</i>' +
    '</p>' +

    '</div>'
  );
}

function createNasStatusBox_(nasStatus) {
  if (!nasStatus.hasAffectedAirports) {
    return (
      '<span style="' +
      'display:inline-block;' +
      'min-width:72px;' +
      'box-sizing:border-box;' +
      'text-align:center;' +
      'padding:7px 12px;' +
      'border:1px solid #238636;' +
      'border-radius:5px;' +
      'background-color:#d9f2df;' +
      'color:#155724;' +
      'font-family:Arial,sans-serif;' +
      'font-weight:bold;' +
      'font-size:12px;' +
      'line-height:14px;' +
      'white-space:nowrap;' +
      '">' +
      'NONE' +
      '</span>'
    );
  }

  const affectedText = nasStatus.affectedAirports
    .map(function(airport) {
      return airport.id;
    })
    .join(', ');

  return (
    '<span style="' +
    'display:inline-block;' +
    'box-sizing:border-box;' +
    'padding:7px 12px;' +
    'border:1px solid #e67700;' +
    'border-radius:5px;' +
    'background-color:#ffe0b2;' +
    'color:#8a4100;' +
    'font-family:Arial,sans-serif;' +
    'font-weight:bold;' +
    'font-size:12px;' +
    'line-height:14px;' +
    'white-space:normal;' +
    '">' +
    'AFFECTED: ' +
    escapeHtml_(affectedText) +
    '</span>'
  );
}

function createStatusBox_(rating) {
  const normalized = String(rating).toUpperCase();

  let backgroundColor = '#d9f2df';
  let borderColor = '#238636';
  let textColor = '#155724';

  if (normalized.indexOf('WATCH') !== -1) {
    backgroundColor = '#fff3bf';
    borderColor = '#d39e00';
    textColor = '#735c00';
  }

  if (normalized.indexOf('ELEVATED') !== -1) {
    backgroundColor = '#ffe0b2';
    borderColor = '#e67700';
    textColor = '#8a4100';
  }

  return (
    '<span style="' +
    'display:inline-block;' +
    'min-width:72px;' +
    'box-sizing:border-box;' +
    'text-align:center;' +
    'padding:7px 12px;' +
    'border:1px solid ' + borderColor + ';' +
    'border-radius:5px;' +
    'background-color:' + backgroundColor + ';' +
    'color:' + textColor + ';' +
    'font-family:Arial,sans-serif;' +
    'font-weight:bold;' +
    'font-size:12px;' +
    'line-height:14px;' +
    'white-space:nowrap;' +
    '">' +
    escapeHtml_(rating) +
    '</span>'
  );
}

function findNasAirport_(nasStatus, airportId) {
  if (!nasStatus.hasAffectedAirports) {
    return null;
  }

  for (
    let index = 0;
    index < nasStatus.affectedAirports.length;
    index += 1
  ) {
    if (nasStatus.affectedAirports[index].id === airportId) {
      return nasStatus.affectedAirports[index];
    }
  }

  return null;
}

function createPlainTextEmail_(report) {
  const lines = [];

  lines.push('KONT DIVERSION WATCH');

  lines.push(
    'Issued: ' +
    report.issuedLocal +
    ' / ' +
    report.issuedZulu
  );

  lines.push('Overall: ' + report.overallRating);

  if (report.nasStatus.hasAffectedAirports) {
    lines.push(
      'NAS Status: AFFECTED — ' +
      report.nasStatus.affectedAirports
        .map(function(airport) {
          return airport.id + ': ' + airport.events.join(' | ');
        })
        .join(' — ')
    );
  } else {
    lines.push('NAS Status: NONE');

    if (report.nasStatus.retrievalMessage) {
      lines.push(report.nasStatus.retrievalMessage);
    }
  }

  lines.push('');
  lines.push('SOURCE AIRPORT RISK');

  report.sourceAirports.forEach(function(airport) {
    const nasAirport = findNasAirport_(
      report.nasStatus,
      airport.id
    );

    lines.push(
      airport.id +
      ': ' +
      airport.rating +
      ' — ' +
      airport.drivers.join('; ')
    );

    if (nasAirport) {
      lines.push(
        'FAA NAS: ' +
        nasAirport.events.join(' | ')
      );
    }
  });

  lines.push('');
  lines.push('LINKS');
  lines.push(CONFIG.awcLink);
  lines.push(CONFIG.nasLink);
  lines.push(CONFIG.delayLink);
  lines.push(CONFIG.notamLink);
  lines.push('');
  lines.push('AIRPORT DETAIL');

  report.allAirports.forEach(function(airport) {
    const nasAirport = findNasAirport_(
      report.nasStatus,
      airport.id
    );

    lines.push(airport.id + ' — ' + airport.rating);
    lines.push(
      'Drivers: ' +
      airport.drivers.join('; ')
    );

    if (nasAirport) {
      lines.push(
        'FAA NAS: ' +
        nasAirport.events.join(' | ')
      );
    }

    lines.push('METAR: ' + airport.metar);
    lines.push('TAF: ' + airport.taf);
    lines.push('');
  });

  lines.push(
    'Decision support only—not a Part 121 dispatch release. ' +
    'Verify FAA flow, NOTAMs, radar/lightning, ADS-B behavior, ' +
    'and KONT operational capacity before action.'
  );

  return lines.join('\n');
}

function saveReportToLog_(report) {
  const spreadsheet = getOrCreateSpreadsheet_();

  const logSheet = spreadsheet.getSheetByName(
    CONFIG.logSheetName
  );

  const nasLogText = report.nasStatus.hasAffectedAirports
    ? report.nasStatus.displayText
    : 'NONE';

  report.allAirports.forEach(function(airport) {
    logSheet.appendRow([
      report.issuedLocal,
      report.issuedZulu,
      report.overallRating,
      nasLogText,
      airport.id,
      airport.rating,
      airport.drivers.join('; '),
      airport.metar,
      airport.taf
    ]);
  });
}

function removeExistingTrigger_(functionName) {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function unique_(items) {
  return items.filter(function(item, index) {
    return items.indexOf(item) === index;
  });
}

function escapeRegex_(value) {
  return String(value).replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&'
  );
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
