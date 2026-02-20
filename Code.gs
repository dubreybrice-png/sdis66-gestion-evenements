/****************************************************
 * SDIS 66 - Gestion des Événements & Candidatures
 * Version: v1.1 | 2026-02-13
 ****************************************************/

const SS_ID = "19aTCFsHGl3NVOvG98-xUhXZ3aAIoVDvCCoIr8FJ3Y2w";
const SHEET_EVENTS  = "Feuille 1";
const SHEET_LISTING = "Listing";
const SHEET_SCORING = "Scoring";
const ADMIN_PWD     = "Sdis66!";

/* ========== WEBAPP ========== */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile("Index");
  template.deepLinkEventId = (e && e.parameter && e.parameter.candidater) ? e.parameter.candidater : '';
  return template.evaluate()
    .setTitle("SDIS 66 - Gestion Événements")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, minimum-scale=0.5, maximum-scale=5, user-scalable=yes');
}

/* ========== ADMIN ========== */
function verifyAdmin(pwd) {
  return pwd === ADMIN_PWD;
}

/* ========== AGENTS (Listing) ========== */
function getAgentsList() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_LISTING);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var data = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      nom: String(data[i][0]).trim(),
      email: String(data[i][1]).trim(),
      matricule: String(data[i][2]).trim(),
      notif: data[i][3] === true || String(data[i][3]).toUpperCase() === "TRUE"
    });
  }
  return result;
}

/* ========== ÉVÉNEMENTS ========== */

/** Récupère les événements à venir (date >= aujourd'hui) */
function getUpcomingEvents() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var events = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;

    var dateEvt = new Date(data[i][2]);
    dateEvt.setHours(0, 0, 0, 0);
    // Le lendemain de l'événement, on le cache
    if (dateEvt < today) continue;

    var candidatsStr = String(data[i][8] || "");
    var retenusStr   = String(data[i][9] || "");
    var candidats = candidatsStr ? candidatsStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    var retenus   = retenusStr   ? retenusStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];

    events.push({
      id: String(data[i][0]),
      nom: String(data[i][1]),
      date: formatDate_(dateEvt),
      dateRaw: dateEvt.getTime(),
      heureDebut: formatTime_(data[i][3]),
      heureFin: formatTime_(data[i][4]),
      lieu: String(data[i][5]),
      commentaire: String(data[i][6] || ""),
      places: Number(data[i][7]),
      candidats: candidats,
      retenus: retenus,
      statut: String(data[i][10] || "ouvert"),
      type: String(data[i][11] || "Autres")
    });
  }

  // Tri par date croissante
  events.sort(function(a, b) { return a.dateRaw - b.dateRaw; });
  return events;
}

/** Crée un événement + notifie les abonnés */
function createEvent(data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);

  // Init en-têtes si nécessaire
  if (!sh.getRange("A1").getValue()) {
    sh.getRange("A1:L1").setValues([["ID","Nom","Date","Heure début","Heure fin","Lieu","Commentaire","Places","Candidats","Retenus","Statut","Type"]]);
    sh.getRange("A1:L1").setFontWeight("bold");
  }

  var dateObj = new Date(data.date + "T12:00:00");

  // Duplicate check: same type + nom + date + heureDebut + heureFin + lieu
  var lastRowCheck = sh.getLastRow();
  if (lastRowCheck >= 2) {
    var existing = sh.getRange(2, 2, lastRowCheck - 1, 11).getValues();
    for (var j = 0; j < existing.length; j++) {
      var exDate = existing[j][1];
      var exDateStr = "";
      if (exDate instanceof Date) {
        exDateStr = exDate.getFullYear() + "-" + String(exDate.getMonth()+1).padStart(2,"0") + "-" + String(exDate.getDate()).padStart(2,"0");
      }
      var exHeureDebut = formatTime_(existing[j][2]);
      var exHeureFin = formatTime_(existing[j][3]);
      if (String(existing[j][0]).trim() === String(data.nom).trim() &&
          exDateStr === data.date &&
          exHeureDebut === String(data.heureDebut).trim() &&
          exHeureFin === String(data.heureFin).trim() &&
          String(existing[j][4]).trim() === String(data.lieu).trim() &&
          String(existing[j][10]).trim() === String(data.type || "Autres").trim()) {
        return { success: false, message: "Un événement identique existe déjà (même nom, date, horaires, lieu et type). Modifiez au moins un élément." };
      }
    }
  }

  var id = Utilities.getUuid().substring(0, 8);

  sh.appendRow([
    id, data.nom, dateObj, data.heureDebut, data.heureFin,
    data.lieu, data.commentaire || "", Number(data.places),
    "", "", "ouvert", data.type || "Autres"
  ]);

  var lastRow = sh.getLastRow();
  sh.getRange(lastRow, 3).setNumberFormat("dd/MM/yyyy");

  // Notification aux abonnés
  try {
    notifyNewEvent_({
      id: id,
      nom: data.nom,
      date: formatDate_(dateObj),
      heureDebut: data.heureDebut,
      heureFin: data.heureFin,
      lieu: data.lieu,
      commentaire: data.commentaire || ""
    });
  } catch (e) {
    Logger.log("Erreur notification nouveau evt: " + e);
  }

  return { success: true, id: id };
}

/** Modifie un événement existant */
function updateEvent(data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var row = findEventRow_(ss, data.id);
  if (!row) return { success: false, message: "Événement non trouvé" };

  var dateObj = new Date(data.date + "T12:00:00");
  sh.getRange(row, 2, 1, 7).setValues([[
    data.nom, dateObj, data.heureDebut, data.heureFin,
    data.lieu, data.commentaire || "", Number(data.places)
  ]]);
  sh.getRange(row, 3).setNumberFormat("dd/MM/yyyy");
  sh.getRange(row, 12).setValue(data.type || "Autres");

  return { success: true };
}

/** Supprime un événement */
function deleteEvent(id) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var row = findEventRow_(ss, id);
  if (!row) return { success: false, message: "Événement non trouvé" };

  sh.deleteRow(row);
  updateScoring();
  return { success: true };
}

/* ========== CANDIDATURES ========== */

/** Un agent postule à un événement */
function postuler(eventId, agentNom) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: false, message: "Aucun événement" };

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== String(eventId)) continue;

    // Vérifier si clôturé
    if (String(data[i][10]) === "cloturé") {
      return { success: false, message: "Cet événement est clôturé, vous ne pouvez plus candidater." };
    }

    // Vérifier doublon
    var candidatsStr = String(data[i][8] || "");
    var candidats = candidatsStr ? candidatsStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    if (candidats.indexOf(agentNom) !== -1) {
      return { success: false, message: "Vous avez déjà candidaté pour cet événement." };
    }

    candidats.push(agentNom);
    sh.getRange(i + 2, 9).setValue(candidats.join(", "));

    updateScoring();
    return { success: true, message: "Candidature enregistrée !" };
  }
  return { success: false, message: "Événement non trouvé" };
}

/** Admin sélectionne les candidats retenus → emails + clôture */
function selectCandidats(eventId, retenusList) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: false };

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== String(eventId)) continue;

    var rowNum = i + 2;
    sh.getRange(rowNum, 10).setValue(retenusList.join(", "));
    sh.getRange(rowNum, 11).setValue("cloturé");

    // Détails de l'événement
    var evt = {
      nom: String(data[i][1]),
      date: formatDate_(new Date(data[i][2])),
      heureDebut: formatTime_(data[i][3]),
      heureFin: formatTime_(data[i][4]),
      lieu: String(data[i][5]),
      commentaire: String(data[i][6] || "")
    };

    // Tous les candidats & non retenus
    var candidatsStr = String(data[i][8] || "");
    var allCandidats = candidatsStr ? candidatsStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    var nonRetenus = allCandidats.filter(function(n) { return retenusList.indexOf(n) === -1; });

    // Map email
    var agents = getAgentsList();
    var emailMap = {};
    agents.forEach(function(a) { emailMap[a.nom] = a.email; });

    var commentPart = evt.commentaire ? " (" + evt.commentaire + ")" : "";

    // Mails retenus
    retenusList.forEach(function(nom) {
      if (emailMap[nom]) {
        try {
          MailApp.sendEmail({
            to: emailMap[nom],
            subject: "SDIS 66 - Vous avez été retenu(e) : " + evt.nom,
            body: "Bonjour,\n\nVous avez été retenu(e) pour l'événement \"" + evt.nom +
                  "\" du " + evt.date + " à " + evt.heureDebut + " jusqu'à " + evt.heureFin +
                  " qui aura lieu " + formatLieu_(evt.lieu) + "." + commentPart +
                  "\n\nCordialement,\nSDIS 66"
          });
        } catch (e) { Logger.log("Mail retenu erreur (" + nom + "): " + e); }
      }
    });

    // Mails non retenus
    nonRetenus.forEach(function(nom) {
      if (emailMap[nom]) {
        try {
          MailApp.sendEmail({
            to: emailMap[nom],
            subject: "SDIS 66 - Candidature non retenue : " + evt.nom,
            body: "Bonjour,\n\nNous avons le regret de vous informer que vous n'avez pas été retenu(e) pour l'événement \"" +
                  evt.nom + "\" du " + evt.date + " à " + evt.heureDebut + " jusqu'à " + evt.heureFin +
                  " qui aura lieu " + formatLieu_(evt.lieu) + "." + commentPart +
                  "\n\nCordialement,\nSDIS 66"
          });
        } catch (e) { Logger.log("Mail non retenu erreur (" + nom + "): " + e); }
      }
    });

    updateScoring();
    return { success: true, message: "Sélection validée, notifications envoyées." };
  }
  return { success: false, message: "Événement non trouvé" };
}

/* ========== NOTIFICATIONS ========== */

/** Inscription aux notifications (vérifie le matricule) */
function subscribeNotif(nom, matricule) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_LISTING);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: false, message: "Agent non trouvé" };

  var data = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === nom) {
      if (String(data[i][2]).trim() !== String(matricule).trim()) {
        return { success: false, message: "Matricule incorrect." };
      }
      sh.getRange(i + 2, 4).setValue(true);
      return { success: true, message: "Inscription réussie ! Vous serez notifié(e) par mail à chaque création d'événement." };
    }
  }
  return { success: false, message: "Agent non trouvé dans le listing." };
}

/** Envoie un mail aux abonnés lors d'un nouvel événement */
function notifyNewEvent_(evt) {
  var agents = getAgentsList();
  var subscribers = agents.filter(function(a) { return a.notif && a.email; });

  // Build direct link to candidature
  var webAppUrl = ScriptApp.getService().getUrl();
  var candidateLink = webAppUrl + (webAppUrl.indexOf('?') === -1 ? '?' : '&') + 'candidater=' + (evt.id || '');

  var lieuDisplay = formatLieu_(evt.lieu);
  var heureFinPart = evt.heureFin ? ' - ' + evt.heureFin : '';
  var commentPart = evt.commentaire ? '\n• Commentaire : ' + evt.commentaire : '';

  var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#1565c0;color:white;padding:20px 24px;border-radius:12px 12px 0 0;">' +
    '<h2 style="margin:0;font-size:1.2rem;">🚒 SDIS 66 — Nouvel événement</h2></div>' +
    '<div style="padding:24px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 12px 12px;">' +
    '<h3 style="color:#1565c0;margin-top:0;">' + evt.nom + '</h3>' +
    '<p style="line-height:1.8;color:#333;">'+
    '📅 <strong>' + evt.date + '</strong><br>' +
    '🕐 ' + evt.heureDebut + heureFinPart + '<br>' +
    '📍 ' + lieuDisplay +
    (evt.commentaire ? '<br>💬 <em>' + evt.commentaire + '</em>' : '') +
    '</p>' +
    '<div style="text-align:center;margin:28px 0 16px;">' +
    '<a href="' + candidateLink + '" style="display:inline-block;background:#43a047;color:white;' +
    'padding:14px 32px;border-radius:10px;text-decoration:none;font-weight:bold;font-size:1rem;' +
    'box-shadow:0 3px 8px rgba(67,160,71,0.3);">📝 Cliquez ici pour candidater !</a></div>' +
    '<p style="font-size:0.85rem;color:#888;text-align:center;">Ou rendez-vous sur l\'application : ' +
    '<a href="' + webAppUrl + '">' + webAppUrl + '</a></p>' +
    '<hr style="border:none;border-top:1px solid #eee;margin:20px 0;">' +
    '<p style="font-size:0.8rem;color:#aaa;">Cordialement,<br>SDIS 66 — Gestion Événements</p>' +
    '</div></div>';

  var textBody = 'Bonjour,\n\nUn nouvel événement a été créé :\n\n' +
    '• ' + evt.nom + '\n' +
    '• Date : ' + evt.date + '\n' +
    '• Horaires : ' + evt.heureDebut + heureFinPart + '\n' +
    '• Lieu : ' + lieuDisplay + '\n' +
    commentPart +
    '\n👉 Cliquez ici pour candidater : ' + candidateLink +
    '\n\nOu rendez-vous sur l\'application : ' + webAppUrl +
    '\n\nCordialement,\nSDIS 66';

  subscribers.forEach(function(agent) {
    try {
      MailApp.sendEmail({
        to: agent.email,
        subject: '🚒 SDIS 66 - Nouvel événement : ' + evt.nom,
        body: textBody,
        htmlBody: htmlBody
      });
    } catch (e) { Logger.log('Notif new event erreur (' + agent.nom + '): ' + e); }
  });
}

/* ========== SCORING ========== */

/** Met à jour l'onglet Scoring : candidatures, sélections, taux, tri décroissant */
function updateScoring() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var listSh = ss.getSheetByName(SHEET_LISTING);
  var evSh   = ss.getSheetByName(SHEET_EVENTS);
  var scSh   = ss.getSheetByName(SHEET_SCORING);

  // En-têtes scoring
  if (!scSh.getRange("A1").getValue()) {
    scSh.getRange("A1:D1").setValues([["Identité","Candidatures","Sélections","Taux de sélection"]]);
    scSh.getRange("A1:D1").setFontWeight("bold");
  }

  // Liste des agents
  var listLast = listSh.getLastRow();
  if (listLast < 2) return;
  var agents = [];
  var listData = listSh.getRange(2, 1, listLast - 1, 1).getValues();
  for (var j = 0; j < listData.length; j++) {
    var n = String(listData[j][0]).trim();
    if (n) agents.push(n);
  }

  // Compteurs
  var candidatures = {};
  var selections = {};
  agents.forEach(function(n) { candidatures[n] = 0; selections[n] = 0; });

  var evLast = evSh.getLastRow();
  if (evLast >= 2) {
    var evData = evSh.getRange(2, 9, evLast - 1, 2).getValues(); // colonnes I et J
    for (var k = 0; k < evData.length; k++) {
      var cands = evData[k][0] ? String(evData[k][0]).split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
      var rets  = evData[k][1] ? String(evData[k][1]).split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
      cands.forEach(function(n) { if (candidatures.hasOwnProperty(n)) candidatures[n]++; });
      rets.forEach(function(n)  { if (selections.hasOwnProperty(n))  selections[n]++;  });
    }
  }

  // Construire les données triées par candidatures décroissant
  var scoringData = agents.map(function(nom) {
    var c = candidatures[nom] || 0;
    var s = selections[nom] || 0;
    var taux = c > 0 ? Math.round((s / c) * 100) + "%" : "–";
    return [nom, c, s, taux];
  });
  scoringData.sort(function(a, b) { return b[1] - a[1]; });

  // Écriture
  var scLast = scSh.getLastRow();
  if (scLast > 1) scSh.getRange(2, 1, scLast - 1, 4).clearContent();
  if (scoringData.length > 0) {
    scSh.getRange(2, 1, scoringData.length, 4).setValues(scoringData);
  }
}

/* ========== ALERTES AUTOMATIQUES ========== */

/** Vérifie les événements et envoie des alertes si nécessaire (à lancer via trigger quotidien) */
function checkAlertsAndSendEmails() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  var now = new Date();
  var alerts = [];

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    var dateEvt = new Date(data[i][2]);
    dateEvt.setHours(0, 0, 0, 0);
    
    var nowMs = now.getTime();
    var evtMs = dateEvt.getTime();
    var diff = evtMs - nowMs;
    var diffHours = diff / (1000 * 60 * 60);
    
    var candidatsStr = String(data[i][8] || "");
    var retenusStr = String(data[i][9] || "");
    var statut = String(data[i][10] || "ouvert");
    
    var candidats = candidatsStr ? candidatsStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    var retenus = retenusStr ? retenusStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    
    var nom = String(data[i][1]);
    var heureDebut = formatTime_(data[i][3]);
    var heureFin = formatTime_(data[i][4]);
    var lieu = String(data[i][5]);
    var type = String(data[i][11] || "Autres");
    
    // Alerte 1 : Événement dans 48h sans aucun candidat
    if (diffHours > 0 && diffHours <= 48 && candidats.length === 0 && statut !== "cloturé") {
      alerts.push({
        type: "48h_no_candidat",
        nom: nom,
        date: formatDate_(dateEvt),
        heureDebut: heureDebut,
        heureFin: heureFin,
        lieu: lieu,
        typeEvt: type,
        delai: Math.round(diffHours) + "h"
      });
    }
    
    // Alerte 2 : Événement dans 24h avec candidats mais pas de sélection
    if (diffHours > 0 && diffHours <= 24 && candidats.length > 0 && retenus.length === 0 && statut !== "cloturé") {
      alerts.push({
        type: "24h_no_selection",
        nom: nom,
        date: formatDate_(dateEvt),
        heureDebut: heureDebut,
        heureFin: heureFin,
        lieu: lieu,
        typeEvt: type,
        candidats: candidats.join(", "),
        delai: Math.round(diffHours) + "h"
      });
    }
  }
  
  if (alerts.length === 0) return;
  
  // Envoi d'un mail récapitulatif
  var recipients = "brice.dubrey@sdis66.fr,florian.bois@sdis66.fr";
  var subject = "⚠️ SDIS 66 - Alertes événements (" + alerts.length + " alerte(s))";
  
  var body = "Bonjour,\n\nVoici les alertes concernant les événements à venir :\n\n";
  
  alerts.forEach(function(alert, idx) {
    body += "─────────────────────────────────────\n";
    body += "ALERTE #" + (idx + 1) + " : ";
    
    if (alert.type === "48h_no_candidat") {
      body += "Aucun candidat (délai: " + alert.delai + ")\n";
      body += "• Événement : " + alert.nom + "\n";
      body += "• Type : " + alert.typeEvt + "\n";
      body += "• Date : " + alert.date + " | " + alert.heureDebut + " - " + alert.heureFin + "\n";
      body += "• Lieu : " + alert.lieu + "\n";
      body += "⚠️ Aucun candidat ne s'est inscrit pour cet événement qui a lieu dans moins de 48h.\n";
    } else if (alert.type === "24h_no_selection") {
      body += "Sélection en attente (délai: " + alert.delai + ")\n";
      body += "• Événement : " + alert.nom + "\n";
      body += "• Type : " + alert.typeEvt + "\n";
      body += "• Date : " + alert.date + " | " + alert.heureDebut + " - " + alert.heureFin + "\n";
      body += "• Lieu : " + alert.lieu + "\n";
      body += "• Candidats : " + alert.candidats + "\n";
      body += "⚠️ Des candidats sont inscrits mais aucune sélection n'a été effectuée. L'événement a lieu dans moins de 24h.\n";
    }
    body += "\n";
  });
  
  body += "─────────────────────────────────────\n";
  body += "Merci de traiter ces alertes rapidement.\n\nCordialement,\nSDIS 66 - Gestion Événements";
  
  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      body: body
    });
    Logger.log("Alertes envoyées : " + alerts.length);
  } catch (e) {
    Logger.log("Erreur envoi alertes : " + e);
  }
}

/* ========== CANDIDATE SCORING ========== */

/** Calcule les scores de candidature pour tous les agents.
 *  Score = +1 par candidature. Quand sélectionné (et pas seul candidat) → score remis à 0. */
function getCandidateScores() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var evSh = ss.getSheetByName(SHEET_EVENTS);
  var evLast = evSh.getLastRow();
  if (evLast < 2) return {};

  var evData = evSh.getRange(2, 1, evLast - 1, 12).getValues();

  // Sort chronologically by date
  evData.sort(function(a, b) {
    var da = a[2] instanceof Date ? a[2].getTime() : 0;
    var db = b[2] instanceof Date ? b[2].getTime() : 0;
    return da - db;
  });

  var scores = {};

  for (var i = 0; i < evData.length; i++) {
    var candidatsStr = String(evData[i][8] || "");
    var retenusStr = String(evData[i][9] || "");
    var candidats = candidatsStr ? candidatsStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    var retenus = retenusStr ? retenusStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];

    // +1 point per candidature
    candidats.forEach(function(nom) {
      if (!scores[nom]) scores[nom] = { postulations: 0, selections: 0, score: 0 };
      scores[nom].postulations++;
      scores[nom].score++;
    });

    // Selected → lose all points, UNLESS only candidate
    retenus.forEach(function(nom) {
      if (!scores[nom]) scores[nom] = { postulations: 0, selections: 0, score: 0 };
      scores[nom].selections++;
      if (candidats.length > 1) {
        scores[nom].score = 0;
      }
    });
  }

  return scores;
}

/* ========== SCRAPER HELPERS ========== */

/** Vérifie si des événements existent déjà (pour le scraper)
 *  Match sur : date + heureDebut + lieu */
function checkEventsExistence(checkList) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();

  if (lastRow < 2) return checkList.map(function() { return false; });

  var data = sh.getRange(2, 2, lastRow - 1, 11).getValues();

  return checkList.map(function(item) {
    for (var j = 0; j < data.length; j++) {
      var exDate = data[j][1];
      var exDateStr = "";
      if (exDate instanceof Date) {
        exDateStr = exDate.getFullYear() + "-" + String(exDate.getMonth()+1).padStart(2,"0") + "-" + String(exDate.getDate()).padStart(2,"0");
      }
      var exHeureDebut = formatTime_(data[j][2]);
      var exLieu = String(data[j][4]).trim();
      var exType = String(data[j][10] || "").trim();

      // Match on date + heureDebut + lieu
      if (exDateStr === item.date &&
          exHeureDebut === String(item.heureDebut || "").trim() &&
          exLieu === String(item.lieu || "").trim() &&
          exType === String(item.type || "SSO ICP").trim()) {
        return true;
      }
    }
    return false;
  });
}

/** Crée plusieurs événements en batch (pour le scraper) */
function createEventsBatch(eventsList) {
  var created = 0;
  var errors = [];

  for (var i = 0; i < eventsList.length; i++) {
    try {
      var r = createEvent(eventsList[i]);
      if (r.success) {
        created++;
      } else {
        errors.push(eventsList[i].nom + " " + eventsList[i].date + ": " + (r.message || "erreur"));
      }
    } catch (e) {
      errors.push(eventsList[i].nom + ": " + e.message);
    }
  }

  return {
    success: true,
    created: created,
    total: eventsList.length,
    errors: errors
  };
}

/* ========== HELPERS ========== */

/** Formate le lieu pour affichage (enlève le préfixe "Autre: ") */
function formatLieu_(lieu) {
  if (lieu && lieu.indexOf("Autre: ") === 0) return lieu.substring(7);
  return lieu || "";
}

function findEventRow_(ss, id) {
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return null;
}

function formatDate_(d) {
  return String(d.getDate()).padStart(2, "0") + "/" +
         String(d.getMonth() + 1).padStart(2, "0") + "/" +
         d.getFullYear();
}

function formatTime_(val) {
  if (val instanceof Date) {
    return String(val.getHours()).padStart(2, "0") + ":" + String(val.getMinutes()).padStart(2, "0");
  }
  return String(val || "");
}
