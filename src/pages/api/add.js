import fs from "fs";
import path from "path";

// Next.js doit accepter un JSON un peu plus gros que la limite par défaut
// (1mb) puisqu'on envoie les fichiers Excel encodés en base64.
export const config = {
  api: {
    bodyParser: {
      sizeLimit: "15mb",
    },
  },
};

const DOSSIERS = {
  moltaqa: "Moltaqa",
  social: "Social",
};

function sanitizeFileName(name) {
  return String(name || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function extensionDepuisNom(nomFichier) {
  return nomFichier && /\.xls$/i.test(nomFichier) ? ".xls" : ".xlsx";
}

function enregistrerFichier(publicDir, dossier, nomParrain, fichier) {
  if (!fichier || !fichier.data) return false;

  const dirAbsolu = path.join(publicDir, dossier);
  fs.mkdirSync(dirAbsolu, { recursive: true });

  const ext = extensionDepuisNom(fichier.nomFichier);
  const cheminFichier = path.join(dirAbsolu, `${nomParrain}${ext}`);

  fs.writeFileSync(cheminFichier, Buffer.from(fichier.data, "base64"));
  return true;
}

export default function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", ["POST"]);
    return res.status(405).json({ error: "Méthode non autorisée." });
  }

  try {
    const { nom, moltaqa, social } = req.body || {};

    const nomPropre = sanitizeFileName(nom);
    if (!nomPropre) {
      return res.status(400).json({ error: "Le nom du parrain est requis." });
    }

    if (!moltaqa?.data && !social?.data) {
      return res
        .status(400)
        .json({ error: "Ajoutez au moins un fichier (Moltaqa ou Social)." });
    }

    const publicDir = path.join(process.cwd(), "public");

    const moltaqaEnregistre = enregistrerFichier(
      publicDir,
      DOSSIERS.moltaqa,
      nomPropre,
      moltaqa
    );
    const socialEnregistre = enregistrerFichier(
      publicDir,
      DOSSIERS.social,
      nomPropre,
      social
    );

    return res.status(200).json({
      ok: true,
      nom: nomPropre,
      moltaqa: moltaqaEnregistre,
      social: socialEnregistre,
    });
  } catch (err) {
    console.error("Erreur /api/excels/add:", err);
    return res
      .status(500)
      .json({ error: "Erreur serveur lors de l'ajout du parrain." });
  }
}
