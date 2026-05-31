/****************************************************
 * SDIS 66 - Gestion des Événements & Candidatures
 * Version: v1.1 | 2026-02-13
 ****************************************************/

const SS_ID = "19aTCFsHGl3NVOvG98-xUhXZ3aAIoVDvCCoIr8FJ3Y2w";
const SHEET_EVENTS  = "Feuille 1";
const SHEET_LISTING = "Listing";
const SHEET_SCORING = "Scoring";
const ADMIN_PWD     = "0007";

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

/**
 * Envoi manuel de notification : liste les événements sans candidats à brice.dubrey@sdis66.fr (mode test).
 */
function sendNotificationManuelle(pwd) {
  if (pwd !== ADMIN_PWD) return 'Mot de passe incorrect.';

  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 'Aucun événement trouvé.';

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);

  var eventsToNotify = [];
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][0]);
    if (!id) continue;
    var dateEvt = new Date(data[i][2]); dateEvt.setHours(0, 0, 0, 0);
    if (dateEvt < today) continue; // événement passé
    var statut = String(data[i][7] || '');
    if (statut === 'cloturé') continue;
    var candidatsRaw = String(data[i][8] || '');
    var candidats = candidatsRaw.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    if (candidats.length > 0) continue; // déjà des candidats
    eventsToNotify.push({
      nom: String(data[i][1]),
      date: formatDate_(new Date(data[i][2])),
      lieu: String(data[i][5])
    });
  }

  if (eventsToNotify.length === 0) return 'Aucun événement sans inscrits à notifier.';

  var webAppUrl = ScriptApp.getService().getUrl();
  var s = eventsToNotify.length > 1 ? 's' : '';

  var htmlBody = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' +
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    '<h2 style="color:#1565c0;">SDIS 66 - ' + eventsToNotify.length + ' evenement' + s + ' sans inscrit' + s + '</h2>' +
    '<p>Bonjour,<br>Les evenements suivants n\'ont pas encore de candidat :</p>';

  var textBody = 'Bonjour,\n\nLes evenements suivants n\'ont pas encore de candidat :\n\n';

  for (var j = 0; j < eventsToNotify.length; j++) {
    var evt = eventsToNotify[j];
    htmlBody += '<p style="border-left:4px solid #1565c0;padding-left:10px;margin:12px 0;">' +
      '<strong>' + evt.nom + '</strong><br>' +
      'Date : ' + evt.date + '<br>' +
      'Lieu : ' + evt.lieu + '</p>';
    textBody += '- ' + evt.nom + ' | ' + evt.date + ' | ' + evt.lieu + '\n';
  }

  htmlBody += '<p><a href="' + webAppUrl + '" style="background:#1565c0;color:white;padding:10px 20px;text-decoration:none;border-radius:4px;font-weight:bold;">Acceder a l\'application</a></p>' +
    '<p style="color:#888;font-size:0.85rem;">Cordialement,<br>SDIS 66 - Gestion Evenements</p>' +
    '</div></body></html>';

  textBody += '\nApplication : ' + webAppUrl + '\n\nCordialement,\nSDIS 66';

  var agents = getAgentsList();
  var subscribers = agents.filter(function(a) { return a.notif && a.email; });
  if (subscribers.length === 0) return 'Aucun abonne trouve dans le Listing.';

  var subject = 'SDIS 66 - ' + eventsToNotify.length + ' evenement' + s + ' sans inscrit' + s;
  var errors = 0;
  subscribers.forEach(function(agent) {
    try {
      GmailApp.sendEmail(agent.email, subject, textBody, { htmlBody: htmlBody, name: 'SDIS 66 - Gestion Evenements' });
    } catch(e) { errors++; Logger.log('Erreur envoi ' + agent.email + ' : ' + e); }
  });

  return 'Mail envoye a ' + (subscribers.length - errors) + '/' + subscribers.length + ' abonnes (' + eventsToNotify.length + ' evenement' + s + ').';
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

  // Ajouter l'événement à la file d'attente pour le digest quotidien (1h du matin)
  try {
    var props = PropertiesService.getScriptProperties();
    var pending = props.getProperty('PENDING_NEW_EVENTS');
    var list = pending ? JSON.parse(pending) : [];
    list.push(id);
    props.setProperty('PENDING_NEW_EVENTS', JSON.stringify(list));
  } catch (e) {
    Logger.log('Erreur ajout file digest: ' + e);
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

/** Admin retire un candidat d'un événement */
function removeCandidat(eventId, agentNom) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: false, message: "Aucun événement" };

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== String(eventId)) continue;

    var candidatsStr = String(data[i][8] || "");
    var candidats = candidatsStr ? candidatsStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    var idx = candidats.indexOf(agentNom);
    if (idx === -1) return { success: false, message: "Cet agent n'est pas candidat." };

    candidats.splice(idx, 1);
    sh.getRange(i + 2, 9).setValue(candidats.join(", "));

    // Si l'agent était aussi retenu, le retirer des retenus
    var retenusStr = String(data[i][9] || "");
    var retenus = retenusStr ? retenusStr.split(",").map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    var rIdx = retenus.indexOf(agentNom);
    if (rIdx !== -1) {
      retenus.splice(rIdx, 1);
      sh.getRange(i + 2, 10).setValue(retenus.join(", "));
    }

    // Si plus aucun candidat ni retenu, remettre le statut à "ouvert"
    if (candidats.length === 0 && retenus.length <= 0) {
      sh.getRange(i + 2, 10).setValue("");
      sh.getRange(i + 2, 11).setValue("ouvert");
    }

    updateScoring();
    return { success: true, message: "Candidat retiré.", remaining: candidats.length };
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
          GmailApp.sendEmail(emailMap[nom],
            "SDIS 66 - Vous avez ete retenu(e) : " + evt.nom,
            "Bonjour,\n\nVous avez ete retenu(e) pour l'evenement \"" + evt.nom +
            "\" du " + evt.date + " a " + evt.heureDebut + " jusqu'a " + evt.heureFin +
            " qui aura lieu " + formatLieu_(evt.lieu) + "." + commentPart +
            "\n\nCordialement,\nSDIS 66",
            { name: 'SDIS 66 - Gestion Evenements' });
        } catch (e) { Logger.log("Mail retenu erreur (" + nom + "): " + e); }
      }
    });

    // Mails non retenus
    nonRetenus.forEach(function(nom) {
      if (emailMap[nom]) {
        try {
          GmailApp.sendEmail(emailMap[nom],
            "SDIS 66 - Candidature non retenue : " + evt.nom,
            "Bonjour,\n\nVotre candidature n'a pas ete retenue pour l'evenement \"" +
            evt.nom + "\" du " + evt.date + " a " + evt.heureDebut + " jusqu'a " + evt.heureFin +
            " qui aura lieu " + formatLieu_(evt.lieu) + "." + commentPart +
            "\n\nCordialement,\nSDIS 66",
            { name: 'SDIS 66 - Gestion Evenements' });
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

/**
 * Digest quotidien à 18h : envoie UN seul mail récapitulatif
 * aux abonnés s'il y a eu de nouveaux événements dans la journée.
 * Si testMode=true, envoie seulement à brice.dubrey@sdis66.fr et auto-supprime le trigger test.
 */
function sendDailyNewEventsDigest(testMode) {
  var props = PropertiesService.getScriptProperties();

  // Auto-suppression du trigger test
  if (testMode) {
    var triggers = ScriptApp.getProjectTriggers();
    for (var t = 0; t < triggers.length; t++) {
      if (triggers[t].getHandlerFunction() === 'sendDailyNewEventsDigestTest') {
        ScriptApp.deleteTrigger(triggers[t]);
      }
    }
  }

  var pending = props.getProperty('PENDING_NEW_EVENTS');
  if (!pending) {
    if (testMode) {
      GmailApp.sendEmail('brice.dubrey@sdis66.fr', '[TEST] Digest SDIS66 - aucun evenement en attente',
        'Le systeme fonctionne correctement. Aucun evenement cree aujourd\'hui.\n\nSi tu crées un événement dans l\'appli, il apparaîtra dans le prochain envoi à 18h.',
        { name: 'SDIS 66 - Gestion Evenements' });
    }
    return;
  }
  var pendingIds = JSON.parse(pending);
  if (!pendingIds || pendingIds.length === 0) return;

  // Vider la file immédiatement
  props.deleteProperty('PENDING_NEW_EVENTS');

  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();

  var events = [];
  for (var i = 0; i < data.length; i++) {
    if (pendingIds.indexOf(String(data[i][0])) === -1) continue;
    var dateEvt = new Date(data[i][2]);
    events.push({
      id: String(data[i][0]),
      nom: String(data[i][1]),
      date: formatDate_(dateEvt),
      heureDebut: formatTime_(data[i][3]),
      heureFin: formatTime_(data[i][4]),
      lieu: String(data[i][5]),
      commentaire: String(data[i][6] || ''),
      type: String(data[i][11] || 'Autres')
    });
  }

  if (events.length === 0) return;

  var agents = getAgentsList();
  var subscribers;
  if (testMode) {
    subscribers = [{ nom: 'Brice (TEST)', email: 'brice.dubrey@sdis66.fr' }];
  } else {
    subscribers = agents.filter(function(a) { return a.notif && a.email; });
  }
  if (subscribers.length === 0) return;

  var webAppUrl = ScriptApp.getService().getUrl();
  var count = events.length;
  var s = count > 1 ? 's' : '';
  var testBannerHtml = testMode ? '<p style="background:#fff3e0;border:1px solid #ff9800;padding:8px 12px;"><strong>[MAIL DE TEST]</strong> - Envoi reel a 18h a tous les abonnes.</p>' : '';

  var htmlBody = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' +
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    testBannerHtml +
    '<h2 style="color:#1565c0;">SDIS 66 - ' + count + ' nouvel' + s + ' evenement' + s + ' a pourvoir</h2>';

  var textBody = (testMode ? '[MAIL DE TEST - envoi reel a 18h]\n\n' : '') +
    'Bonjour,\n\n' + count + ' nouvel' + s + ' evenement' + s + ' ' + (count > 1 ? 'ont' : 'a') + ' ete cree' + s + ' aujourd\'hui :\n\n';

  for (var j = 0; j < events.length; j++) {
    var evt = events[j];
    var lieuDisplay = formatLieu_(evt.lieu);
    var heureFinPart = evt.heureFin ? ' - ' + evt.heureFin : '';
    var candidateLink = webAppUrl + (webAppUrl.indexOf('?') === -1 ? '?' : '&') + 'candidater=' + evt.id;

    htmlBody += '<hr>' +
      '<h3 style="color:#1565c0;margin-bottom:4px;">' + evt.nom + '</h3>' +
      '<p style="margin:4px 0;">Date : <strong>' + evt.date + '</strong><br>' +
      'Horaires : ' + evt.heureDebut + heureFinPart + '<br>' +
      'Lieu : ' + lieuDisplay +
      (evt.commentaire ? '<br>Commentaire : <em>' + evt.commentaire + '</em>' : '') +
      '</p>' +
      '<p><a href="' + candidateLink + '" style="background:#43a047;color:white;padding:8px 18px;text-decoration:none;border-radius:4px;font-weight:bold;">Candidater</a></p>';

    textBody += evt.nom + '\n' +
      'Date : ' + evt.date + '\n' +
      'Horaires : ' + evt.heureDebut + heureFinPart + '\n' +
      'Lieu : ' + lieuDisplay + '\n' +
      (evt.commentaire ? 'Commentaire : ' + evt.commentaire + '\n' : '') +
      'Candidater : ' + candidateLink + '\n\n';
  }

  htmlBody += '<hr><p><a href="' + webAppUrl + '" style="color:#1565c0;">Acceder a l\'application</a></p>' +
    '<p style="color:#888;font-size:0.85rem;">Cordialement,<br>SDIS 66 - Gestion Evenements</p>' +
    '</div></body></html>';

  textBody += 'Application : ' + webAppUrl + '\n\nCordialement,\nSDIS 66';

  var subject = (testMode ? '[TEST] ' : '') + 'SDIS 66 - ' + count + ' nouvel' + s + ' evenement' + s + ' a pourvoir';

  subscribers.forEach(function(agent) {
    try {
      GmailApp.sendEmail(agent.email, subject, textBody, { htmlBody: htmlBody, name: 'SDIS 66 - Gestion Evenements' });
    } catch (e) { Logger.log('Digest erreur (' + agent.nom + '): ' + e); }
  });

  Logger.log('Digest envoyé : ' + count + ' événement(s) à ' + subscribers.length + ' abonné(s)' + (testMode ? ' [TEST]' : ''));
}

/**
 * Wrapper appelé par le trigger de test (une seule fois).
 * Peut aussi être appelé directement depuis l'éditeur pour tester immédiatement.
 */
function sendDailyNewEventsDigestTest() {
  sendDailyNewEventsDigest(true);
}

/**
 * Installe un trigger one-shot à 11h35 pour tester le digest.
 * Auto-supprime le trigger après le premier envoi.
 */
function setupTestTrigger() {
  // Supprimer un éventuel ancien trigger test
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'sendDailyNewEventsDigestTest') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  // Créer trigger à 14h58 aujourd'hui (ou demain si l'heure est passée)
  var d = new Date();
  d.setHours(15, 1, 0, 0);
  if (d <= new Date()) {
    d.setDate(d.getDate() + 1);
  }
  ScriptApp.newTrigger('sendDailyNewEventsDigestTest')
    .timeBased()
    .at(d)
    .create();
  Logger.log('Trigger test installé pour : ' + d.toLocaleString());
  return 'Trigger test installé pour ' + d.toLocaleString('fr-FR');
}

/**
 * Installe les triggers automatiques :
 * - Digest nouveaux événements : tous les jours à 1h
 * - Alertes 24h/48h : tous les jours à 7h (existant)
 * À exécuter UNE FOIS manuellement.
 */
function setupTriggers() {
  // Supprimer les anciens triggers
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    var fn = existing[i].getHandlerFunction();
    if (fn === 'sendDailyNewEventsDigest' || fn === 'notifyNewEvent_' || fn === 'checkAlertsAndSendEmails') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }

  // Digest quotidien à 18h
  ScriptApp.newTrigger('sendDailyNewEventsDigest')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  // Alertes 24h/48h à 7h du matin
  ScriptApp.newTrigger('checkAlertsAndSendEmails')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  return 'Triggers installés : digest 18h + alertes 7h';
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
    GmailApp.sendEmail(recipients, subject, body, { name: 'SDIS 66 - Gestion Evenements' });
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

/** Retourne les données de scoring de tous les agents pour affichage */
function getScoresData() {
  var agents = getAgentsList();
  var scores = getCandidateScores();
  var result = [];
  for (var i = 0; i < agents.length; i++) {
    var nom = agents[i].nom;
    var s = scores[nom] || { postulations: 0, selections: 0, score: 0 };
    var ratio = s.postulations > 0 ? Math.round(s.selections / s.postulations * 100) : 0;
    result.push({
      nom: nom,
      postulations: s.postulations,
      selections: s.selections,
      ratio: ratio,
      score: s.score
    });
  }
  result.sort(function(a, b) { return b.postulations - a.postulations; });
  return result;
}

/* ========== FMPA (consultation) ========== */

/** Récupère les prochaines sessions FMPA depuis le spreadsheet Confirmation */
function getFmpaData() {
  var FMPA_SS_ID = "1hmLYGOcu0tt1y4Cg9GIquGM1QVua3Hi37pcacPZrW4Q";
  var FMPA_SHEET = "Confirmation";

  var ss = SpreadsheetApp.openById(FMPA_SS_ID);
  var sh = ss.getSheetByName(FMPA_SHEET);
  if (!sh) return [];

  var lastCol = sh.getLastColumn();
  if (lastCol < 2) return []; // col A = labels, données à partir de col B

  // Ligne 1 = dates (texte: "06/01/2026", "12/01/2026 M", "03/02/2026 AM", "17/01/2026 Matin")
  // Ligne 2 = thème (XABCDE, SSO + PISU, SSO + Carrefour Technique)
  // Lignes 3-15 = apprenants
  // Lignes 18-21 = formateurs
  var dates      = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]; // getDisplayValues pour avoir le texte exact
  var titres     = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
  var apprenants = sh.getRange(3, 1, 13, lastCol).getValues(); // lignes 3 à 15
  var formateurs = sh.getRange(18, 1, 4, lastCol).getValues(); // lignes 18 à 21

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var sessions = [];

  // Commencer à col index 1 (col B) car col A = labels
  for (var col = 1; col < lastCol; col++) {
    var dateStr = String(dates[col] || "").trim();
    var titreVal = String(titres[col] || "").trim();

    if (!dateStr) continue;

    // Parser la date depuis le texte: "06/01/2026" ou "12/01/2026 M" ou "03/02/2026 AM" ou "17/01/2026 Matin"
    var dateMatch = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (!dateMatch) continue;

    var day = parseInt(dateMatch[1], 10);
    var month = parseInt(dateMatch[2], 10) - 1; // mois 0-indexé
    var year = parseInt(dateMatch[3], 10);
    var dateObj = new Date(year, month, day);
    dateObj.setHours(0, 0, 0, 0);

    // Ignorer les FMPA passées
    if (dateObj < today) continue;

    // Déterminer les horaires depuis la ligne 1 (suffixe après la date)
    var horaires = "";
    var afterDate = dateStr.replace(/\d{2}\/\d{2}\/\d{4}/, "").trim();
    if (/^AM$/i.test(afterDate)) {
      horaires = "13h30 – 17h30";
    } else if (/^M$|^Matin$/i.test(afterDate)) {
      horaires = "8h30 – 12h30";
    }

    // Déterminer le type FMPA depuis ligne 2
    var typeFmpa = "";
    if (/xabcde/i.test(titreVal)) {
      typeFmpa = "FMPA 1";
      if (!horaires) horaires = "8h30 – 17h30"; // journée entière par défaut
    } else if (/sso.*pisu|pisu.*sso/i.test(titreVal)) {
      typeFmpa = "FMPA 2";
    } else if (/sso.*carrefour|carrefour.*technique/i.test(titreVal)) {
      typeFmpa = "FMPA 2 (Carrefour Technique)";
    } else {
      typeFmpa = titreVal || "FMPA";
    }

    // Horaires depuis le thème ligne 2 (ex: "SSO + PISU M" = matin)
    if (!horaires && /\bM\b/.test(titreVal) && !/Matin/i.test(afterDate)) {
      horaires = "8h30 – 12h30";
    }

    // Récupérer les noms des apprenants (lignes 3-15)
    var noms = [];
    for (var r = 0; r < 13; r++) {
      var nom = String(apprenants[r][col] || "").trim();
      if (nom) noms.push(nom);
    }

    // Récupérer les formateurs (lignes 18-21)
    var formList = [];
    for (var r = 0; r < 4; r++) {
      var form = String(formateurs[r][col] || "").trim();
      if (form) formList.push(form);
    }

    sessions.push({
      date: formatDate_(dateObj),
      dateRaw: dateObj.getTime(),
      type: typeFmpa,
      horaires: horaires,
      apprenants: noms,
      formateurs: formList
    });
  }

  // Tri par date croissante
  sessions.sort(function(a, b) { return a.dateRaw - b.dateRaw; });

  return sessions;
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

/** DEBUG: list all upcoming events with lieu + type */
function debugListEvents() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  var today = new Date(); today.setHours(0,0,0,0);
  var result = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var dateEvt = new Date(data[i][2]); dateEvt.setHours(0,0,0,0);
    if (dateEvt < today) continue;
    result.push({
      row: i+2,
      id: String(data[i][0]),
      nom: String(data[i][1]),
      date: formatDate_(dateEvt),
      lieu: String(data[i][5]),
      type: String(data[i][11] || ""),
      typeLen: String(data[i][11] || "").length,
      typeCharCodes: String(data[i][11] || "").split('').map(function(c){ return c.charCodeAt(0); }).join(',')
    });
  }
  return result;
}
/* ============================================================
   BILAN SEMAINE — statistiques 20-26 avril 2026
   ============================================================ */
function getBilanSemaine() {
  var now = new Date();
  var day = now.getDay();
  var diffToMon = (day === 0 ? -6 : 1 - day);
  var mon = new Date(now); mon.setDate(now.getDate() + diffToMon); mon.setHours(0,0,0,0);
  var sun = new Date(mon); sun.setDate(mon.getDate() + 6); sun.setHours(23,59,59,999);
  var START = mon, END = sun;
  var YEAR_START = new Date(START.getFullYear(), 0, 1);
  var fmt = function(d){ return d.toLocaleDateString('fr-FR', {day:'numeric',month:'long'}); };

  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_EVENTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { error: 'Aucune donnée' };

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  // Récupérer les grades depuis le Listing pour détecter les ISP
  var shListing = ss.getSheetByName(SHEET_LISTING);
  var listingLastRow = shListing.getLastRow();
  var listingData = listingLastRow >= 2 ? shListing.getRange(2, 1, listingLastRow - 1, 4).getValues() : [];
  // Construire un set de noms ISP (col 3 = index 3 = grade/type)
  // Si pas de colonne grade, on garde tous les noms pour reference
  var allAgentNames = {};
  for (var li = 0; li < listingData.length; li++) {
    if (listingData[li][0]) allAgentNames[String(listingData[li][0]).trim()] = true;
  }

  var evenementsSemaine = 0;
  var evenementsAvecInscrit = 0;
  var evenementsAvecIsp = 0; // type contient ISP ou au moins un candidat inscrit
  var totalAnnee = 0;
  var typesCounts = {};

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var dateEvt = new Date(data[i][2]);
    if (isNaN(dateEvt.getTime())) continue;
    dateEvt.setHours(0,0,0,0);
    var type = String(data[i][11] || 'Autres');
    var candidatsStr = String(data[i][8] || '');
    var candidats = candidatsStr ? candidatsStr.split(',').map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
    var nbInscrits = candidats.length;

    if (dateEvt >= YEAR_START) {
      totalAnnee++;
    }
    if (dateEvt >= START && dateEvt <= END) {
      evenementsSemaine++;
      if (nbInscrits > 0) evenementsAvecInscrit++;
      // ISP = type contient "ISP" ou "isp" ou "infirmier"
      if (/isp|infirm/i.test(type) || /isp|infirm/i.test(String(data[i][1]))) {
        evenementsAvecIsp++;
      }
      typesCounts[type] = (typesCounts[type] || 0) + 1;
    }
  }

  var result = {
    periode: fmt(START) + ' au ' + fmt(END),
    evenementsSemaine: evenementsSemaine,
    evenementsAvecInscrit: evenementsAvecInscrit,
    evenementsAvecIsp: evenementsAvecIsp,
    totalAnnee2026: totalAnnee,
    typesCounts: typesCounts
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}