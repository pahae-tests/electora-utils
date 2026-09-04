import fs from "fs";
import path from "path";

const DOSSIERS = ["Moltaqa", "Social"];

export default function handler(req, res) {
  if (req.method !== "POST" && req.method !== "DELETE") {
    res.setHeader("Allow", ["POST", "DELETE"]);
    return res.status(405).json({ error: "Méthode non autorisée." });
  }

  try {
    const { nom } = req.body || {};
    if (!nom || typeof nom !== "string") {
      return res.status(400).json({ error: "Le nom du parrain est requis." });
    }

    // path.basename empêche toute tentative de sortir du dossier (ex: "../../etc").
    const nomSecurise = path.basename(nom).trim();
    if (!nomSecurise) {
      return res.status(400).json({ error: "Nom de parrain invalide." });
    }

    const publicDir = path.join(process.cwd(), "public");
    let fichiersSupprimes = 0;

    DOSSIERS.forEach((dossier) => {
      const dossierAbsolu = path.join(publicDir, dossier);
      if (!fs.existsSync(dossierAbsolu)) return;

      const fichiers = fs
        .readdirSync(dossierAbsolu)
        .filter((f) => f.replace(/\.xlsx?$/i, "") === nomSecurise);

      fichiers.forEach((f) => {
        fs.unlinkSync(path.join(dossierAbsolu, f));
        fichiersSupprimes += 1;
      });
    });

    if (fichiersSupprimes === 0) {
      return res.status(404).json({ error: "Aucun fichier trouvé pour ce parrain." });
    }

    return res.status(200).json({ ok: true, fichiersSupprimes });
  } catch (err) {
    console.error("Erreur /api/excels/delete:", err);
    return res
      .status(500)
      .json({ error: "Erreur serveur lors de la suppression." });
  }
}
