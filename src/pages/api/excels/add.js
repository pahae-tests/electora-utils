import fs from "fs";
import path from "path";

// Les fichiers Excel peuvent être assez volumineux lorsqu'ils
// sont envoyés en base64.
export const config = {
  api: {
    bodyParser: {
      sizeLimit: "100mb",
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

function extensionValide(nomFichier) {
  return /\.xlsx?$/i.test(nomFichier || "");
}

function enregistrerFichier(publicDir, dossier, fichier) {
  if (!fichier || !fichier.data || !fichier.nomFichier) {
    return false;
  }

  const nomFichier = sanitizeFileName(fichier.nomFichier);

  if (!nomFichier || !extensionValide(nomFichier)) {
    return false;
  }

  const dirAbsolu = path.join(publicDir, dossier);

  fs.mkdirSync(dirAbsolu, { recursive: true });

  const cheminFichier = path.join(dirAbsolu, nomFichier);

  fs.writeFileSync(
    cheminFichier,
    Buffer.from(fichier.data, "base64")
  );

  return true;
}

export default function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", ["POST"]);

    return res.status(405).json({
      error: "Méthode non autorisée.",
    });
  }

  try {
    const { moltaqa = [], social = [] } = req.body || {};

    if (!Array.isArray(moltaqa) || !Array.isArray(social)) {
      return res.status(400).json({
        error: "Format des fichiers invalide.",
      });
    }

    if (moltaqa.length === 0 && social.length === 0) {
      return res.status(400).json({
        error: "Ajoutez au moins un fichier Moltaqa ou Social.",
      });
    }

    const publicDir = path.join(process.cwd(), "public");

    let moltaqaAjoutes = 0;
    let socialAjoutes = 0;

    // =========================
    // FICHIERS MOLTAQA
    // =========================

    moltaqa.forEach((fichier) => {
      if (enregistrerFichier(publicDir, DOSSIERS.moltaqa, fichier)) {
        moltaqaAjoutes++;
      }
    });

    // =========================
    // FICHIERS SOCIAL
    // =========================

    social.forEach((fichier) => {
      if (enregistrerFichier(publicDir, DOSSIERS.social, fichier)) {
        socialAjoutes++;
      }
    });

    return res.status(200).json({
      ok: true,
      moltaqa: moltaqaAjoutes,
      social: socialAjoutes,
      total: moltaqaAjoutes + socialAjoutes,
    });
  } catch (err) {
    console.error("Erreur /api/excels/add:", err);

    return res.status(500).json({
      error: "Erreur serveur lors de l'ajout des fichiers.",
    });
  }
}