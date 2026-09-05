import { useMemo, useState } from "react";
import { useRouter } from "next/router";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";
import {
    FileSpreadsheet,
    Download,
    FolderOpen,
    Plus,
    X,
    Trash2,
    Loader2,
} from "lucide-react";

// =========================================================
// LECTURE DES FICHIERS EXCEL DIRECTEMENT DANS /PUBLIC
// =========================================================

function listerFichiersExcel(publicDir) {
    try {
        return fs
            .readdirSync(publicDir)
            .filter((f) => /\.xlsx?$/i.test(f))
            .sort((a, b) => a.localeCompare(b, "fr"));
    } catch (err) {
        return [];
    }
}

// Un fichier est considéré comme vide s'il n'a aucune donnée exploitable.
function excelEstVide(cheminAbsolu) {
    try {
        const stat = fs.statSync(cheminAbsolu);
        if (stat.size === 0) return true;

        const classeur = XLSX.readFile(cheminAbsolu);

        if (!classeur.SheetNames || classeur.SheetNames.length === 0) {
            return true;
        }

        return classeur.SheetNames.every((nomFeuille) => {
            const feuille = classeur.Sheets[nomFeuille];

            const lignes = XLSX.utils.sheet_to_json(feuille, {
                header: 1,
                blankrows: false,
            });

            return lignes.length <= 1;
        });
    } catch (err) {
        return true;
    }
}

// =========================================================
// SERVER SIDE
// =========================================================

export async function getServerSideProps() {
    const publicDir = path.join(process.cwd(), "public");

    const fichiers = listerFichiersExcel(publicDir);

    const fichiersExcel = fichiers.map((nomFichier) => ({
        nomFichier,
        nomParrain: nomFichier.replace(/\.xlsx?$/i, "").trim(),
        url: `/${nomFichier}`,
        vide: excelEstVide(path.join(publicDir, nomFichier)),
    }));

    return {
        props: {
            fichiersExcel,
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

function sanitizeFileName(name) {
    const cleaned = String(name || "")
        .trim()
        .replace(/[\\/:*?"<>|]/g, "")
        .replace(/\s+/g, " ")
        .trim();

    return cleaned || "export";
}

function estDisponible(fichier) {
    return Boolean(fichier) && !fichier.vide;
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
// MODALE D'AJOUT D'UN FICHIER
// =========================================================

function ModaleAjout({ onClose, onSuccess }) {
    const [fichier, setFichier] = useState(null);
    const [envoiEnCours, setEnvoiEnCours] = useState(false);
    const [erreur, setErreur] = useState("");

    async function handleSubmit(e) {
        e.preventDefault();
        setErreur("");

        if (!fichier) {
            setErreur("Sélectionnez un fichier Excel.");
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
                    reject(
                        new Error("Impossible de lire le fichier.")
                    );
                };

                reader.readAsDataURL(fichier);
            });

            const payload = {
                fichiers: [
                    {
                        nomFichier: fichier.name,
                        data: dataBase64,
                    },
                ],
            };

            const res = await fetch("/api/excels/add", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify(payload),
            });

            const data = await res.json().catch(() => ({}));

            if (!res.ok) {
                throw new Error(
                    data.error || "Une erreur est survenue."
                );
            }

            onSuccess();
        } catch (err) {
            setErreur(
                err.message ||
                    "Une erreur est survenue."
            );
        } finally {
            setEnvoiEnCours(false);
        }
    }

    return (
        <div
            className="modalOverlay"
            onClick={onClose}
        >
            <div
                className="modalBox"
                onClick={(e) => e.stopPropagation()}
            >
                <div className="modalHeader">
                    <h2>Ajouter un fichier Excel</h2>

                    <button
                        type="button"
                        className="modalClose"
                        onClick={onClose}
                        aria-label="Fermer"
                    >
                        <X
                            size={18}
                            strokeWidth={2}
                        />
                    </button>
                </div>

                <form
                    onSubmit={handleSubmit}
                    className="modalForm"
                >
                    <label className="fieldLabel">
                        Fichier Excel

                        <input
                            type="file"
                            accept=".xlsx,.xls"
                            onChange={(e) =>
                                setFichier(
                                    e.target.files?.[0] || null
                                )
                            }
                        />

                        {fichier && (
                            <div className="selectedFiles">
                                <div className="selectedFile">
                                    <FileSpreadsheet
                                        size={14}
                                    />

                                    <span>
                                        {fichier.name}
                                    </span>
                                </div>
                            </div>
                        )}
                    </label>

                    <p className="modalHint">
                        Le nom du fichier sera utilisé comme nom du parrain.
                    </p>

                    {erreur && (
                        <p className="modalErreur">
                            {erreur}
                        </p>
                    )}

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
                                <Plus
                                    size={14}
                                    strokeWidth={2}
                                />
                            )}

                            {envoiEnCours
                                ? "Ajout..."
                                : "Ajouter le fichier"}
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

export default function ExcelsPage({ fichiersExcel }) {
    const router = useRouter();

    const [search, setSearch] = useState("");
    const [modaleOuverte, setModaleOuverte] = useState(false);
    const [suppressionEnCours, setSuppressionEnCours] =
        useState(null);

    const fichiersFiltres = useMemo(() => {
        const q = normaliser(search);

        if (!q) return fichiersExcel;

        return fichiersExcel.filter((fichier) =>
            normaliser(fichier.nomParrain).includes(q)
        );
    }, [fichiersExcel, search]);

    const totalFichiers = fichiersExcel.length;

    const totalDisponibles = fichiersExcel.filter(
        (fichier) => estDisponible(fichier)
    ).length;

    const totalVides = fichiersExcel.filter(
        (fichier) => fichier.vide
    ).length;

    function rafraichir() {
        router.replace(
            router.asPath,
            undefined,
            { scroll: false }
        );
    }

    async function handleSupprimer(fichier) {
        alert('mzl madrtha') return
        const confirme = window.confirm(
            `Supprimer définitivement le fichier "${fichier.nomFichier}" ?`
        );

        if (!confirme) return;

        setSuppressionEnCours(
            fichier.nomFichier
        );

        try {
            const res = await fetch(
                "/api/excels/delete",
                {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                    },
                    body: JSON.stringify({
                        nomFichier:
                            fichier.nomFichier,
                    }),
                }
            );

            const data =
                await res
                    .json()
                    .catch(() => ({}));

            if (!res.ok) {
                throw new Error(
                    data.error ||
                        "Erreur lors de la suppression."
                );
            }

            rafraichir();
        } catch (err) {
            window.alert(
                err.message ||
                    "Erreur lors de la suppression."
            );
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
                        <span className="eyebrow">
                            Dossiers Excel
                        </span>

                        <h1>
                            Fichiers par parrain
                        </h1>

                        <p>
                            Tous les fichiers Excel présents
                            directement dans le dossier public.
                        </p>
                    </div>
                </header>

                {/* STATS */}
                <section className="statsRow">
                    <div className="statCard">
                        <span className="statValue">
                            {totalFichiers}
                        </span>

                        <span className="statLabel">
                            Fichiers
                        </span>
                    </div>

                    <div className="statCard">
                        <span className="statValue">
                            {totalDisponibles}
                        </span>

                        <span className="statLabel">
                            Disponibles
                        </span>
                    </div>

                    <div className="statCard">
                        <span className="statValue">
                            {totalVides}
                        </span>

                        <span className="statLabel">
                            Vides
                        </span>
                    </div>
                </section>

                {/* RECHERCHE */}
                <section className="searchSection">
                    <div className="searchRow">
                        <div className="searchBox">
                            <input
                                type="text"
                                placeholder="Rechercher un parrain..."
                                value={search}
                                onChange={(e) =>
                                    setSearch(
                                        e.target.value
                                    )
                                }
                            />
                        </div>

                        <button
                            type="button"
                            disabled
                            className="actionBtn actionBtnPrimary addBtn disabled"
                            onClick={() =>
                                setModaleOuverte(true)
                            }
                        >
                            <Plus
                                size={14}
                                strokeWidth={2}
                            />

                            Ajouter un fichier
                        </button>
                    </div>
                </section>

                {/* LISTE */}
                <section className="listSection">
                    {fichiersFiltres.length === 0 ? (
                        <div className="empty">
                            <FolderOpen
                                size={30}
                                strokeWidth={1.6}
                            />

                            <p>
                                {fichiersExcel.length === 0
                                    ? "Aucun fichier Excel trouvé dans le dossier public."
                                    : "Aucun fichier ne correspond à votre recherche."}
                            </p>
                        </div>
                    ) : (
                        <div className="parrainList">
                            {fichiersFiltres.map(
                                (fichier) => {
                                    const disponible =
                                        estDisponible(
                                            fichier
                                        );

                                    const enSuppression =
                                        suppressionEnCours ===
                                        fichier.nomFichier;

                                    return (
                                        <div
                                            key={
                                                fichier.nomFichier
                                            }
                                            className="parrainCard"
                                        >
                                            <div className="parrainInfo">
                                                <div className="parrainIcon">
                                                    <FileSpreadsheet
                                                        size={18}
                                                        strokeWidth={
                                                            1.8
                                                        }
                                                    />
                                                </div>

                                                <div className="parrainText">
                                                    <span className="parrainNom">
                                                        {
                                                            fichier.nomParrain
                                                        }
                                                    </span>

                                                    <div className="parrainBadges">
                                                        <span
                                                            className={`badge ${
                                                                disponible
                                                                    ? "badgeOn"
                                                                    : "badgeEmpty"
                                                            }`}
                                                        >
                                                            {fichier.vide
                                                                ? "Vide"
                                                                : "Excel"}
                                                        </span>

                                                        <span className="fileName">
                                                            {
                                                                fichier.nomFichier
                                                            }
                                                        </span>
                                                    </div>
                                                </div>
                                            </div>

                                            <div className="parrainActions">
                                                <button
                                                    type="button"
                                                    className={`actionBtn ${
                                                        !disponible
                                                            ? "actionBtnDisabled"
                                                            : ""
                                                    }`}
                                                    disabled={
                                                        !disponible
                                                    }
                                                    onClick={() =>
                                                        disponible &&
                                                        telechargerUnFichier(
                                                            fichier.url,
                                                            fichier.nomFichier
                                                        )
                                                    }
                                                >
                                                    <Download
                                                        size={
                                                            14
                                                        }
                                                        strokeWidth={
                                                            2
                                                        }
                                                    />

                                                    Télécharger
                                                </button>

                                                <button
                                                    type="button"
                                                    disabled
                                                    className="actionBtn actionBtnDanger disabled"
                                                    disabled={
                                                        enSuppression
                                                    }
                                                    onClick={() =>
                                                        handleSupprimer(
                                                            fichier
                                                        )
                                                    }
                                                    aria-label={`Supprimer ${fichier.nomFichier}`}
                                                >
                                                    {enSuppression ? (
                                                        <Loader2
                                                            size={
                                                                14
                                                            }
                                                            strokeWidth={
                                                                2
                                                            }
                                                            className="spin"
                                                        />
                                                    ) : (
                                                        <Trash2
                                                            size={
                                                                14
                                                            }
                                                            strokeWidth={
                                                                2
                                                            }
                                                        />
                                                    )}
                                                </button>
                                            </div>
                                        </div>
                                    );
                                }
                            )}
                        </div>
                    )}
                </section>
            </div>

            {modaleOuverte && (
                <ModaleAjout
                    onClose={() =>
                        setModaleOuverte(false)
                    }
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
                    font-family: Georgia, "Times New Roman",
                        serif;
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
                    font-family: Georgia, "Times New Roman",
                        serif;
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
                }

                .searchBox {
                    position: relative;
                    display: flex;
                    align-items: center;
                    flex: 1;
                    min-width: 0;
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
                    box-shadow: 0 0 0 3px
                        rgba(31, 58, 95, 0.12);
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

                .parrainList {
                    display: flex;
                    flex-direction: column;
                    gap: 10px;
                }

                .parrainCard {
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

                .parrainInfo {
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    min-width: 0;
                }

                .parrainIcon {
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

                .parrainText {
                    display: flex;
                    flex-direction: column;
                    gap: 6px;
                    min-width: 0;
                }

                .parrainNom {
                    font-size: 14px;
                    font-weight: 600;
                    color: #16191c;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }

                .parrainBadges {
                    display: flex;
                    gap: 6px;
                    align-items: center;
                    min-width: 0;
                }

                .badge {
                    font-size: 10px;
                    font-weight: 700;
                    letter-spacing: 0.03em;
                    padding: 2px 8px;
                    border-radius: 3px;
                    text-transform: uppercase;
                }

                .badgeOn {
                    background: #eef2f6;
                    color: #1f3a5f;
                }

                .badgeEmpty {
                    background: #fbf1e4;
                    color: #a3742f;
                }

                .fileName {
                    font-size: 11px;
                    color: #8a8378;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    white-space: nowrap;
                    max-width: 300px;
                }

                .parrainActions {
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

                .actionBtn:hover:not(
                        .actionBtnDisabled
                    ):not(:disabled) {
                    background: #eef2f6;
                }

                .actionBtnDisabled {
                    color: #b3ac9f;
                    border-color: #e5e1d8;
                    cursor: not-allowed;
                    pointer-events: none;
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
                        grid-template-columns: repeat(
                            2,
                            1fr
                        );
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

                    .addBtn {
                        justify-content: center;
                    }

                    .parrainCard {
                        flex-direction: column;
                        align-items: stretch;
                    }

                    .parrainActions {
                        width: 100%;
                        display: grid;
                        grid-template-columns: 1fr auto;
                    }

                    .actionBtn {
                        width: 100%;
                    }

                    .fileName {
                        max-width: 180px;
                    }
                }

                .disabled {
                    cursor: not-allowed;
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
