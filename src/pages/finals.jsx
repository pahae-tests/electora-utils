import { useMemo, useState } from "react";
import { useRouter } from "next/router";
import fs from "fs";
import path from "path";
import {
    FileText,
    Download,
    FolderOpen,
    Plus,
    X,
    Trash2,
    Loader2,
    User,
    Calendar,
} from "lucide-react";

// =========================================================
// LECTURE DES DONNÉES (data/partage.json)
// =========================================================

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

// =========================================================
// SERVER SIDE
// =========================================================

export async function getServerSideProps() {
    const cheminJson = path.join(process.cwd(), "data", "partage.json");

    const documents = lireDocuments(cheminJson).sort((a, b) =>
        String(b.date).localeCompare(String(a.date))
    );

    return {
        props: {
            documentsInitiaux: documents,
        },
    };
}

// =========================================================
// UTILITAIRES
// =========================================================

function normaliser(txt) {
    return String(txt ?? "")
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
}

function formaterDate(dateIso) {
    try {
        const d = new Date(dateIso);
        return d.toLocaleDateString("fr-FR", {
            day: "2-digit",
            month: "long",
            year: "numeric",
        });
    } catch (err) {
        return dateIso;
    }
}

function estAujourdhui(dateIso) {
    const aujourdhui = new Date().toISOString().slice(0, 10);
    return String(dateIso).slice(0, 10) === aujourdhui;
}

// =========================================================
// TÉLÉCHARGEMENT
// =========================================================

async function telechargerUnFichier(url, nomFichier) {
    const response = await fetch(url);

    if (!response.ok) {
        throw new Error("Impossible de télécharger le fichier.");
    }

    const blob = await response.blob();
    const blobUrl = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = nomFichier;
    document.body.appendChild(a);
    a.click();
    a.remove();

    URL.revokeObjectURL(blobUrl);
}

// =========================================================
// MODALE D'AJOUT D'UN DOCUMENT
// =========================================================

function ModaleAjout({ onClose, onSuccess }) {
    const [titre, setTitre] = useState("");
    const [auteur, setAuteur] = useState("");
    const [fichier, setFichier] = useState(null);
    const [envoiEnCours, setEnvoiEnCours] = useState(false);
    const [erreur, setErreur] = useState("");

    async function handleSubmit(e) {
        e.preventDefault();
        setErreur("");

        if (!titre.trim()) {
            setErreur("Précisez un titre.");
            return;
        }

        if (!auteur.trim()) {
            setErreur("Précisez votre nom.");
            return;
        }

        if (!fichier) {
            setErreur("Sélectionnez un document.");
            return;
        }

        setEnvoiEnCours(true);

        try {
            const reader = new FileReader();

            const dataBase64 = await new Promise((resolve, reject) => {
                reader.onload = () => {
                    const resultat = String(reader.result || "");
                    resolve(resultat.split(",")[1] || "");
                };

                reader.onerror = () => {
                    reject(new Error("Impossible de lire le fichier."));
                };

                reader.readAsDataURL(fichier);
            });

            const payload = {
                titre: titre.trim(),
                auteur: auteur.trim(),
                nomFichierOriginal: fichier.name,
                data: dataBase64,
            };

            const res = await fetch("/api/partage/add", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify(payload),
            });

            const data = await res.json().catch(() => ({}));

            if (!res.ok) {
                throw new Error(data.error || "Une erreur est survenue.");
            }

            onSuccess();
        } catch (err) {
            setErreur(err.message || "Une erreur est survenue.");
        } finally {
            setEnvoiEnCours(false);
        }
    }

    return (
        <div className="modalOverlay" onClick={onClose}>
            <div className="modalBox" onClick={(e) => e.stopPropagation()}>
                <div className="modalHeader">
                    <h2>Partager un document</h2>

                    <button
                        type="button"
                        className="modalClose"
                        onClick={onClose}
                        aria-label="Fermer"
                    >
                        <X size={18} strokeWidth={2} />
                    </button>
                </div>

                <form onSubmit={handleSubmit} className="modalForm">
                    <label className="fieldLabel">
                        Titre du document
                        <input
                            type="text"
                            placeholder="Ex : Compte-rendu réunion"
                            value={titre}
                            onChange={(e) => setTitre(e.target.value)}
                            className="fieldInput"
                        />
                    </label>

                    <label className="fieldLabel">
                        Votre nom
                        <input
                            type="text"
                            placeholder="Ex : Yassine"
                            value={auteur}
                            onChange={(e) => setAuteur(e.target.value)}
                            className="fieldInput"
                        />
                    </label>

                    <label className="fieldLabel">
                        Document

                        <input
                            type="file"
                            onChange={(e) =>
                                setFichier(e.target.files?.[0] || null)
                            }
                        />

                        {fichier && (
                            <div className="selectedFiles">
                                <div className="selectedFile">
                                    <FileText size={14} />
                                    <span>{fichier.name}</span>
                                </div>
                            </div>
                        )}
                    </label>

                    <p className="modalHint">
                        La date d'ajout sera enregistrée automatiquement
                        (aujourd'hui).
                    </p>

                    {erreur && <p className="modalErreur">{erreur}</p>}

                    <div className="modalActions">
                        <button
                            type="button"
                            className="actionBtn"
                            onClick={onClose}
                            disabled={envoiEnCours}
                        >
                            Annuler
                        </button>

                        <button
                            type="submit"
                            className="actionBtn actionBtnPrimary"
                            disabled={envoiEnCours}
                        >
                            {envoiEnCours ? (
                                <Loader2
                                    size={14}
                                    strokeWidth={2}
                                    className="spin"
                                />
                            ) : (
                                <Plus size={14} strokeWidth={2} />
                            )}

                            {envoiEnCours ? "Envoi..." : "Partager"}
                        </button>
                    </div>
                </form>
            </div>

            <style jsx>{`
                .modalOverlay {
                    position: fixed;
                    inset: 0;
                    background: rgba(22, 25, 28, 0.45);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    padding: 20px;
                    z-index: 50;
                }

                .modalBox {
                    width: 100%;
                    max-width: 520px;
                    background: #ffffff;
                    border-radius: 8px;
                    border: 1px solid #e0ddd4;
                    max-height: 90vh;
                    overflow-y: auto;
                }

                .modalHeader {
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    padding: 18px 20px;
                    border-bottom: 1px solid #e0ddd4;
                }

                .modalHeader h2 {
                    margin: 0;
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 18px;
                    color: #16191c;
                }

                .modalClose {
                    background: none;
                    border: none;
                    color: #6b6459;
                    cursor: pointer;
                    padding: 4px;
                    display: flex;
                }

                .modalForm {
                    padding: 20px;
                    display: flex;
                    flex-direction: column;
                    gap: 18px;
                }

                .fieldLabel {
                    display: flex;
                    flex-direction: column;
                    gap: 8px;
                    font-size: 12px;
                    font-weight: 600;
                    color: #6b6459;
                    text-transform: uppercase;
                    letter-spacing: 0.03em;
                }

                .fieldLabel input[type="file"] {
                    font-size: 12px;
                    font-weight: 400;
                    text-transform: none;
                    letter-spacing: normal;
                    color: #6b6459;
                }

                .fieldInput {
                    font-size: 13px;
                    font-weight: 400;
                    text-transform: none;
                    letter-spacing: normal;
                    color: #1e2124;
                    background: #fbfaf8;
                    border: 1px solid #d7d2c6;
                    border-radius: 4px;
                    padding: 9px 10px;
                }

                .fieldInput:focus {
                    outline: none;
                    border-color: #1f3a5f;
                    box-shadow: 0 0 0 3px rgba(31, 58, 95, 0.12);
                }

                .selectedFiles {
                    display: flex;
                    flex-direction: column;
                    gap: 5px;
                    padding: 8px;
                    background: #fbfaf8;
                    border: 1px solid #e0ddd4;
                    border-radius: 4px;
                }

                .selectedFile {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    padding: 6px 8px;
                    background: #ffffff;
                    border: 1px solid #e5e1d8;
                    border-radius: 4px;
                    font-size: 12px;
                    color: #1e2124;
                }

                .selectedFile svg {
                    flex-shrink: 0;
                    color: #1f3a5f;
                }

                .selectedFile span {
                    overflow: hidden;
                    text-overflow: ellipsis;
                    white-space: nowrap;
                }

                .modalHint {
                    margin: 0;
                    font-size: 12px;
                    line-height: 1.5;
                    color: #8a8378;
                }

                .modalErreur {
                    margin: 0;
                    font-size: 13px;
                    color: #b3261e;
                    background: #fdecea;
                    border: 1px solid #f3c6c2;
                    border-radius: 4px;
                    padding: 8px 10px;
                }

                .modalActions {
                    display: flex;
                    justify-content: flex-end;
                    gap: 8px;
                    margin-top: 4px;
                }

                .actionBtn {
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                    gap: 6px;
                    background: #ffffff;
                    border: 1px solid #d7d2c6;
                    color: #1f3a5f;
                    border-radius: 4px;
                    padding: 7px 12px;
                    font-size: 12px;
                    font-weight: 600;
                    cursor: pointer;
                    text-decoration: none;
                    font-family: inherit;
                }

                .actionBtn:hover:not(:disabled) {
                    background: #eef2f6;
                }

                .actionBtnPrimary {
                    background: #1f3a5f;
                    border-color: #1f3a5f;
                    color: #ffffff;
                }

                .actionBtnPrimary:hover:not(:disabled) {
                    background: #16293f !important;
                }

                .actionBtn:disabled {
                    opacity: 0.6;
                    cursor: not-allowed;
                }

                .spin {
                    animation: spin 0.8s linear infinite;
                }

                @keyframes spin {
                    from {
                        transform: rotate(0deg);
                    }
                    to {
                        transform: rotate(360deg);
                    }
                }
            `}</style>
        </div>
    );
}

// =========================================================
// PAGE
// =========================================================

export default function FinalsPage({ documentsInitiaux }) {
    const router = useRouter();

    const [documents, setDocuments] = useState(documentsInitiaux || []);
    const [search, setSearch] = useState("");
    const [filtreAuteur, setFiltreAuteur] = useState("tous");
    const [modaleOuverte, setModaleOuverte] = useState(false);
    const [suppressionEnCours, setSuppressionEnCours] = useState(null);

    const auteurs = useMemo(() => {
        const set = new Set(
            documents.map((d) => d.auteur).filter(Boolean)
        );
        return Array.from(set).sort((a, b) => a.localeCompare(b, "fr"));
    }, [documents]);

    const documentsFiltres = useMemo(() => {
        const q = normaliser(search);

        return documents.filter((doc) => {
            const correspondTitre = q
                ? normaliser(doc.titre).includes(q)
                : true;

            const correspondAuteur =
                filtreAuteur === "tous" ? true : doc.auteur === filtreAuteur;

            return correspondTitre && correspondAuteur;
        });
    }, [documents, search, filtreAuteur]);

    const totalDocuments = documents.length;
    const totalAuteurs = auteurs.length;
    const totalAujourdhui = documents.filter((d) =>
        estAujourdhui(d.date)
    ).length;

    function rafraichir() {
        router.replace(router.asPath, undefined, { scroll: false });
    }

    async function handleSupprimer(doc) {
        const confirme = window.confirm(
            `Supprimer définitivement le document "${doc.titre}" ?`
        );

        if (!confirme) return;

        setSuppressionEnCours(doc.id);

        try {
            const res = await fetch("/api/partage/delete", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ id: doc.id }),
            });

            const data = await res.json().catch(() => ({}));

            if (!res.ok) {
                throw new Error(
                    data.error || "Erreur lors de la suppression."
                );
            }

            setDocuments((prev) => prev.filter((d) => d.id !== doc.id));
            rafraichir();
        } catch (err) {
            window.alert(err.message || "Erreur lors de la suppression.");
        } finally {
            setSuppressionEnCours(null);
        }
    }

    return (
        <div className="page">
            <div className="sheet">
                {/* HEADER */}
                <header className="letterhead">
                    <div className="letterheadBar" />

                    <div className="letterheadText">
                        <span className="eyebrow">Espace partagé</span>
                        <h1>Documents</h1>
                        <p>
                            Déposez, retrouvez et téléchargez les documents
                            partagés par l'équipe.
                        </p>
                    </div>
                </header>

                {/* STATS */}
                <section className="statsRow">
                    <div className="statCard">
                        <span className="statValue">{totalDocuments}</span>
                        <span className="statLabel">Documents</span>
                    </div>

                    <div className="statCard">
                        <span className="statValue">{totalAuteurs}</span>
                        <span className="statLabel">Contributeurs</span>
                    </div>

                    <div className="statCard">
                        <span className="statValue">{totalAujourdhui}</span>
                        <span className="statLabel">Ajoutés aujourd'hui</span>
                    </div>
                </section>

                {/* RECHERCHE */}
                <section className="searchSection">
                    <div className="searchRow">
                        <div className="searchBox">
                            <input
                                type="text"
                                placeholder="Rechercher un document..."
                                value={search}
                                onChange={(e) => setSearch(e.target.value)}
                            />
                        </div>

                        <div className="filterBox">
                            <select
                                value={filtreAuteur}
                                onChange={(e) =>
                                    setFiltreAuteur(e.target.value)
                                }
                            >
                                <option value="tous">Tous les auteurs</option>
                                {auteurs.map((a) => (
                                    <option key={a} value={a}>
                                        {a}
                                    </option>
                                ))}
                            </select>
                        </div>

                        <button
                            type="button"
                            className="actionBtn actionBtnPrimary addBtn"
                            onClick={() => setModaleOuverte(true)}
                        >
                            <Plus size={14} strokeWidth={2} />
                            Ajouter un document
                        </button>
                    </div>
                </section>

                {/* LISTE */}
                <section className="listSection">
                    {documentsFiltres.length === 0 ? (
                        <div className="empty">
                            <FolderOpen size={30} strokeWidth={1.6} />
                            <p>
                                {documents.length === 0
                                    ? "Aucun document partagé pour le moment."
                                    : "Aucun document ne correspond à votre recherche."}
                            </p>
                        </div>
                    ) : (
                        <div className="docList">
                            {documentsFiltres.map((doc) => {
                                const enSuppression =
                                    suppressionEnCours === doc.id;

                                return (
                                    <div key={doc.id} className="docCard">
                                        <div className="docInfo">
                                            <div className="docIcon">
                                                <FileText
                                                    size={18}
                                                    strokeWidth={1.8}
                                                />
                                            </div>

                                            <div className="docText">
                                                <span className="docTitre">
                                                    {doc.titre}
                                                </span>

                                                <div className="docMeta">
                                                    <span className="metaItem">
                                                        <User size={11} />
                                                        {doc.auteur}
                                                    </span>

                                                    <span className="metaItem">
                                                        <Calendar size={11} />
                                                        {formaterDate(doc.date)}
                                                    </span>

                                                    <span className="fileName">
                                                        {doc.nomOriginal}
                                                    </span>
                                                </div>
                                            </div>
                                        </div>

                                        <div className="docActions">
                                            <button
                                                type="button"
                                                className="actionBtn"
                                                onClick={() =>
                                                    telechargerUnFichier(
                                                        doc.url,
                                                        doc.nomOriginal
                                                    )
                                                }
                                            >
                                                <Download
                                                    size={14}
                                                    strokeWidth={2}
                                                />
                                                Télécharger
                                            </button>

                                            <button
                                                type="button"
                                                className="actionBtn actionBtnDanger"
                                                disabled={enSuppression}
                                                onClick={() =>
                                                    handleSupprimer(doc)
                                                }
                                                aria-label={`Supprimer ${doc.titre}`}
                                            >
                                                {enSuppression ? (
                                                    <Loader2
                                                        size={14}
                                                        strokeWidth={2}
                                                        className="spin"
                                                    />
                                                ) : (
                                                    <Trash2
                                                        size={14}
                                                        strokeWidth={2}
                                                    />
                                                )}
                                            </button>
                                        </div>
                                    </div>
                                );
                            })}
                        </div>
                    )}
                </section>
            </div>

            {modaleOuverte && (
                <ModaleAjout
                    onClose={() => setModaleOuverte(false)}
                    onSuccess={() => {
                        setModaleOuverte(false);
                        rafraichir();
                    }}
                />
            )}

            <style jsx>{`
                .page {
                    min-height: 100vh;
                    background: #f4f2ee;
                    padding: 56px 20px;
                    display: flex;
                    justify-content: center;
                    font-family: -apple-system, BlinkMacSystemFont,
                        "Segoe UI", Roboto, sans-serif;
                    color: #1e2124;
                }

                .sheet {
                    width: 100%;
                    max-width: 900px;
                    background: #ffffff;
                    border: 1px solid #e0ddd4;
                    border-radius: 4px;
                }

                .letterhead {
                    display: flex;
                    align-items: stretch;
                    border-bottom: 1px solid #e0ddd4;
                }

                .letterheadBar {
                    width: 6px;
                    background: #1f3a5f;
                    flex-shrink: 0;
                }

                .letterheadText {
                    padding: 30px 36px 26px;
                }

                .eyebrow {
                    font-size: 12px;
                    letter-spacing: 0.04em;
                    color: #96723a;
                    font-weight: 600;
                }

                .letterheadText h1 {
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 26px;
                    font-weight: 600;
                    margin: 6px 0 8px;
                    color: #16191c;
                }

                .letterheadText p {
                    margin: 0;
                    font-size: 14px;
                    color: #6b6459;
                }

                .statsRow {
                    display: grid;
                    grid-template-columns: repeat(3, 1fr);
                    gap: 1px;
                    background: #e0ddd4;
                    border-bottom: 1px solid #e0ddd4;
                }

                .statCard {
                    background: #fbfaf8;
                    padding: 16px 20px;
                    display: flex;
                    flex-direction: column;
                    gap: 4px;
                }

                .statValue {
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 22px;
                    font-weight: 600;
                    color: #1f3a5f;
                }

                .statLabel {
                    font-size: 11px;
                    font-weight: 600;
                    letter-spacing: 0.04em;
                    color: #8a8378;
                    text-transform: uppercase;
                }

                .searchSection {
                    padding: 24px 36px 0;
                }

                .searchRow {
                    display: flex;
                    gap: 10px;
                    align-items: stretch;
                    flex-wrap: wrap;
                }

                .searchBox {
                    position: relative;
                    display: flex;
                    align-items: center;
                    flex: 1;
                    min-width: 180px;
                }

                .searchBox input {
                    width: 100%;
                    background: #fbfaf8;
                    border: 1px solid #d7d2c6;
                    color: #1e2124;
                    border-radius: 4px;
                    padding: 10px 12px;
                    font-size: 14px;
                    box-sizing: border-box;
                }

                .searchBox input:focus {
                    outline: none;
                    border-color: #1f3a5f;
                    box-shadow: 0 0 0 3px rgba(31, 58, 95, 0.12);
                }

                .filterBox select {
                    height: 100%;
                    background: #fbfaf8;
                    border: 1px solid #d7d2c6;
                    color: #1e2124;
                    border-radius: 4px;
                    padding: 10px 12px;
                    font-size: 13px;
                    box-sizing: border-box;
                    min-width: 160px;
                }

                .filterBox select:focus {
                    outline: none;
                    border-color: #1f3a5f;
                    box-shadow: 0 0 0 3px rgba(31, 58, 95, 0.12);
                }

                .addBtn {
                    white-space: nowrap;
                    flex-shrink: 0;
                }

                .listSection {
                    padding: 20px 36px 36px;
                }

                .empty {
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    gap: 10px;
                    padding: 50px 20px;
                    color: #9a9488;
                    text-align: center;
                }

                .empty p {
                    margin: 0;
                    font-size: 14px;
                }

                .docList {
                    display: flex;
                    flex-direction: column;
                    gap: 10px;
                }

                .docCard {
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    gap: 16px;
                    padding: 14px 16px;
                    border: 1px solid #e0ddd4;
                    border-radius: 6px;
                    background: #fbfaf8;
                    flex-wrap: wrap;
                }

                .docInfo {
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    min-width: 0;
                }

                .docIcon {
                    width: 36px;
                    height: 36px;
                    min-width: 36px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    background: #ffffff;
                    border: 1px solid #e0ddd4;
                    border-radius: 6px;
                    color: #1f3a5f;
                }

                .docText {
                    display: flex;
                    flex-direction: column;
                    gap: 6px;
                    min-width: 0;
                }

                .docTitre {
                    font-size: 14px;
                    font-weight: 600;
                    color: #16191c;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }

                .docMeta {
                    display: flex;
                    gap: 12px;
                    align-items: center;
                    flex-wrap: wrap;
                    min-width: 0;
                }

                .metaItem {
                    display: inline-flex;
                    align-items: center;
                    gap: 4px;
                    font-size: 11px;
                    color: #8a8378;
                    font-weight: 600;
                }

                .fileName {
                    font-size: 11px;
                    color: #8a8378;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    white-space: nowrap;
                    max-width: 220px;
                }

                .docActions {
                    display: flex;
                    gap: 8px;
                    flex-shrink: 0;
                    flex-wrap: wrap;
                }

                .actionBtn {
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                    gap: 6px;
                    background: #ffffff;
                    border: 1px solid #d7d2c6;
                    color: #1f3a5f;
                    border-radius: 4px;
                    padding: 7px 12px;
                    font-size: 12px;
                    font-weight: 600;
                    cursor: pointer;
                    text-decoration: none;
                    font-family: inherit;
                }

                .actionBtn:hover:not(:disabled) {
                    background: #eef2f6;
                }

                .actionBtnPrimary {
                    background: #1f3a5f;
                    border-color: #1f3a5f;
                    color: #ffffff;
                }

                .actionBtnPrimary:hover:not(:disabled) {
                    background: #16293f !important;
                }

                .actionBtnDanger {
                    color: #b3261e;
                    border-color: #f0d3d1;
                    padding: 7px 9px;
                }

                .actionBtnDanger:hover:not(:disabled) {
                    background: #fdecea;
                }

                .actionBtnDanger:disabled {
                    color: #d9a8a4;
                    cursor: not-allowed;
                }

                .spin {
                    animation: spin 0.8s linear infinite;
                }

                @keyframes spin {
                    from {
                        transform: rotate(0deg);
                    }
                    to {
                        transform: rotate(360deg);
                    }
                }

                @media (max-width: 640px) {
                    .page {
                        padding: 24px 12px;
                    }

                    .letterheadText,
                    .searchSection,
                    .listSection {
                        padding-left: 20px;
                        padding-right: 20px;
                    }

                    .letterheadText {
                        padding-top: 22px;
                        padding-bottom: 20px;
                    }

                    .letterheadText h1 {
                        font-size: 21px;
                    }

                    .letterheadText p {
                        font-size: 13px;
                    }

                    .statsRow {
                        grid-template-columns: repeat(2, 1fr);
                    }

                    .statCard {
                        padding: 12px 14px;
                    }

                    .statCard:last-child {
                        grid-column: 1 / -1;
                    }

                    .searchRow {
                        flex-direction: column;
                    }

                    .filterBox select {
                        width: 100%;
                    }

                    .addBtn {
                        justify-content: center;
                    }

                    .docCard {
                        flex-direction: column;
                        align-items: stretch;
                    }

                    .docActions {
                        width: 100%;
                        display: grid;
                        grid-template-columns: 1fr auto;
                    }

                    .actionBtn {
                        width: 100%;
                    }

                    .fileName {
                        max-width: 150px;
                    }
                }

                @media (max-width: 380px) {
                    .letterheadText h1 {
                        font-size: 19px;
                    }

                    .statValue {
                        font-size: 18px;
                    }

                    .statLabel {
                        font-size: 10px;
                    }
                }
            `}</style>
        </div>
    );
}
