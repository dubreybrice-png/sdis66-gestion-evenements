# SDIS 66 - Gestion des Événements & Candidatures

Application Google Apps Script pour la gestion des événements FMPA, SSUAP, ICP, etc.

## Fonctionnalités

- 📅 Création/modification/suppression d'événements (FMPA, SSUAP, ICP, Autres)
- 👤 Candidature des agents aux événements
- ✅ Sélection des candidats retenus par l'admin
- 📧 Notifications par email (nouveaux événements, résultats de sélection)
- 📊 Scoring : suivi des candidatures et sélections par agent
- ⚠️ Alertes automatiques (48h sans candidat, 24h sans sélection)

## Structure Google Sheets

- **Feuille 1** : Événements (ID, Nom, Date, Heures, Lieu, Commentaire, Places, Candidats, Retenus, Statut, Type)
- **Listing** : Agents (Nom, Email, Matricule, Notif)
- **Scoring** : Stats par agent (Identité, Candidatures, Sélections, Taux)

## Déploiement avec clasp

```bash
# Push vers Google Apps Script
clasp push --force

# Pull depuis Google Apps Script
clasp pull

# Déployer en webapp
clasp deploy
```

## Spreadsheet ID

`19aTCFsHGl3NVOvG98-xUhXZ3aAIoVDvCCoIr8FJ3Y2w`
