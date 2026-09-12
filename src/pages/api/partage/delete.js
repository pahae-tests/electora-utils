import fs from "fs";
import path from "path";

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
        const { id } = req.body || {};

        if (!id) {
            return res.status(400).json({ error: "Identifiant manquant." });
        }

        const dossierPublic = path.join(process.cwd(), "public", "partage");
        const cheminJson = path.join(process.cwd(), "data", "partage.json");

        const documents = lireDocuments(cheminJson);
        const document = documents.find((d) => d.id === id);

        if (!document) {
            return res
                .status(404)
                .json({ error: "Document introuvable." });
        }

        const cheminFichier = path.join(
            dossierPublic,
            document.nomFichier
        );

        if (fs.existsSync(cheminFichier)) {
            fs.unlinkSync(cheminFichier);
        }

        const documentsRestants = documents.filter((d) => d.id !== id);
        ecrireDocuments(cheminJson, documentsRestants);

        return res.status(200).json({ ok: true });
    } catch (err) {
        return res.status(500).json({
            error: "Erreur serveur lors de la suppression.",
        });
    }
}
