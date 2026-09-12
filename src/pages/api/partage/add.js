import fs from "fs";
import path from "path";

export const config = {
    api: {
        bodyParser: {
            sizeLimit: "15mb",
        },
    },
};

function sanitizeFileName(name) {
    const cleaned = String(name || "")
        .trim()
        .replace(/[\\/:*?"<>|]/g, "")
        .replace(/\s+/g, " ")
        .trim();

    return cleaned || "document";
}

function lireDocuments(cheminJson) {
    try {
        if (!fs.existsSync(cheminJson)) return [];
        const contenu = fs.readFileSync(cheminJson, "utf-8");
        const data = JSON.parse(contenu || "[]");
        return Array.isArray(data) ? data : [];
    } catch (err) {
        return [];
    }
}

function ecrireDocuments(cheminJson, documents) {
    fs.writeFileSync(
        cheminJson,
        JSON.stringify(documents, null, 2),
        "utf-8"
    );
}

export default async function handler(req, res) {
    if (req.method !== "POST") {
        res.setHeader("Allow", ["POST"]);
        return res.status(405).json({ error: "Méthode non autorisée." });
    }

    try {
        const { titre, auteur, nomFichierOriginal, data } = req.body || {};

        if (!titre || !String(titre).trim()) {
            return res.status(400).json({ error: "Le titre est requis." });
        }

        if (!auteur || !String(auteur).trim()) {
            return res.status(400).json({ error: "Le nom est requis." });
        }

        if (!nomFichierOriginal || !data) {
            return res
                .status(400)
                .json({ error: "Aucun fichier reçu." });
        }

        const dossierPublic = path.join(process.cwd(), "public", "partage");
        const cheminJson = path.join(process.cwd(), "data", "partage.json");

        if (!fs.existsSync(dossierPublic)) {
            fs.mkdirSync(dossierPublic, { recursive: true });
        }

        if (!fs.existsSync(path.dirname(cheminJson))) {
            fs.mkdirSync(path.dirname(cheminJson), { recursive: true });
        }

        const extension = path.extname(nomFichierOriginal);
        const nomBase = sanitizeFileName(
            path.basename(nomFichierOriginal, extension)
        );

        const id = Date.now().toString();
        const nomFichierStocke = `${id}-${nomBase}${extension}`;
        const cheminFichier = path.join(dossierPublic, nomFichierStocke);

        const buffer = Buffer.from(data, "base64");
        fs.writeFileSync(cheminFichier, buffer);

        const documents = lireDocuments(cheminJson);

        const nouveauDocument = {
            id,
            titre: String(titre).trim(),
            auteur: String(auteur).trim(),
            nomFichier: nomFichierStocke,
            nomOriginal: nomFichierOriginal,
            url: `/partage/${nomFichierStocke}`,
            date: new Date().toISOString().slice(0, 10),
        };

        documents.push(nouveauDocument);
        ecrireDocuments(cheminJson, documents);

        return res.status(200).json({ document: nouveauDocument });
    } catch (err) {
        return res.status(500).json({
            error: "Erreur serveur lors de l'ajout du document.",
        });
    }
}
