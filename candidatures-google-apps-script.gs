// Réception des candidatures : enregistre les dossiers dans Google Drive,
// un dossier par offre, puis un sous-dossier par candidat. Envoie aussi un e-mail.
// À coller dans https://script.google.com (Nouveau projet), puis Déployer > Application Web.

const TOKEN = 'CHANGEZ-CE-CODE';   // doit être identique à CAND_TOKEN dans index.html
const RACINE = 'Candidatures BECP Builder';
const EMAIL = 'becpbuilder@gmail.com';
const MAX_TOTAL = 12 * 1024 * 1024;
const MAX_FICHIERS = 8;
const EXT = /\.(pdf|docx?|jpe?g|png)$/i;

function propre(texte, max) {
  const t = String(texte || '').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
  return t || 'Sans titre';
}

function sousDossier(parent, nom) {
  const it = parent.getFoldersByName(nom);
  return it.hasNext() ? it.next() : parent.createFolder(nom);
}

function dossierRacine() {
  const it = DriveApp.getFoldersByName(RACINE);
  return it.hasNext() ? it.next() : DriveApp.createFolder(RACINE);
}

function reponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return reponse({ ok: false, erreur: 'acces' });
    if (!d.nom || !d.offre || !/^\S+@\S+\.\S+$/.test(d.email || '')) return reponse({ ok: false, erreur: 'champs' });

    const fichiers = Array.isArray(d.fichiers) ? d.fichiers : [];
    let total = 0;
    fichiers.forEach(function (f) { total += String(f.data || '').length * 0.75; });
    if (fichiers.length > MAX_FICHIERS || total > MAX_TOTAL) return reponse({ ok: false, erreur: 'taille' });

    const dOffre = sousDossier(dossierRacine(), propre(d.offre, 80));
    const date = Utilities.formatDate(new Date(), 'Africa/Kinshasa', 'yyyy-MM-dd HH.mm');
    const dCand = sousDossier(dOffre, date + ' - ' + propre(d.nom, 60));

    fichiers.forEach(function (f, i) {
      if (!EXT.test(String(f.nom || ''))) return;
      const role = f.role === 'cv' ? 'CV' : f.role === 'lettre' ? 'Lettre de motivation' : 'Autre ' + (i + 1);
      const nom = role + ' - ' + propre(f.nom, 80);
      dCand.createFile(Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/octet-stream', nom));
    });

    const infos = 'Offre : ' + d.offre + (d.reference ? ' (' + d.reference + ')' : '') +
      '\nNom : ' + d.nom + '\nE-mail : ' + d.email + '\nTéléphone : ' + (d.tel || '') +
      '\n\nMessage :\n' + (d.message || '');
    dCand.createFile('Informations.txt', infos, MimeType.PLAIN_TEXT);

    MailApp.sendEmail(EMAIL, 'Nouvelle candidature : ' + d.offre,
      infos + '\n\nDossier : ' + dCand.getUrl(), { replyTo: d.email });

    return reponse({ ok: true });
  } catch (err) {
    return reponse({ ok: false, erreur: 'serveur' });
  }
}
