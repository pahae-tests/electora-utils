import { useMemo, useState } from "react";
import { useRouter } from "next/router";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";
import {
    FileSpreadsheet,
    Download,
    Search,
    FolderOpen,
    Plus,
    X,
    Trash2,
    Loader2,
} from "lucide-react";

// =========================================================
// LECTURE DES DOSSIERS PUBLIC/MOLTAQA ET PUBLIC/SOCIAL
// =========================================================

const DOSSIERS = [
    { key: "moltaqa", label: "Moltaqa", dir: "Moltaqa" },
    { key: "social", label: "Social", dir: "Social" },
];

function listerFichiersExcel(dirAbsolu) {
    try {
        return fs
            .readdirSync(dirAbsolu)
            .filter((f) => /\.xlsx?$/i.test(f))
            .sort((a, b) => a.localeCompare(b, "fr"));
    } catch (err) {
        // Le dossier n'existe pas encore : on ne casse pas la page.
        return [];
    }
}

// Un fichier est considéré comme "vide" s'il pèse 0 octet, s'il n'a
// aucune feuille, ou si aucune de ses feuilles ne contient de ligne de
// données (au-delà d'un éventuel en-tête).
function excelEstVide(cheminAbsolu) {
    try {
        const stat = fs.statSync(cheminAbsolu);
        if (stat.size === 0) return true;

        const classeur = XLSX.readFile(cheminAbsolu);
        if (!classeur.SheetNames || classeur.SheetNames.length === 0) return true;

        return classeur.SheetNames.every((nomFeuille) => {
            const feuille = classeur.Sheets[nomFeuille];
            const lignes = XLSX.utils.sheet_to_json(feuille, {
                header: 1,
                blankrows: false,
            });
            // 0 ligne = feuille totalement vide. 1 seule ligne = probablement
            // juste la ligne d'en-têtes, sans aucune donnée exploitable.
            return lignes.length <= 1;
        });
    } catch (err) {
        // Fichier illisible/corrompu : on le traite comme indisponible.
        return true;
    }
}

export async function getServerSideProps() {
    const publicDir = path.join(process.cwd(), "public");

    const fichiersParDossier = {};
    DOSSIERS.forEach(({ key, dir }) => {
        fichiersParDossier[key] = listerFichiersExcel(path.join(publicDir, dir));
    });

    // Fusionne les deux listes par nom de parrain (nom du fichier sans extension).
    const parrainsMap = new Map();

    DOSSIERS.forEach(({ key, dir }) => {
        fichiersParDossier[key].forEach((fichier) => {
            const nomParrain = fichier.replace(/\.xlsx?$/i, "").trim();
            const cheminAbsolu = path.join(publicDir, dir, fichier);

            if (!parrainsMap.has(nomParrain)) {
                parrainsMap.set(nomParrain, { nom: nomParrain });
            }

            parrainsMap.get(nomParrain)[key] = {
                url: `/${dir}/${fichier}`,
                vide: excelEstVide(cheminAbsolu),
            };
        });
    });

    const parrains = Array.from(parrainsMap.values()).sort((a, b) =>
        a.nom.localeCompare(b.nom, "fr")
    );

    return { props: { parrains } };
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

// Nom du fichier téléchargé : "<parrain>_<moltaqa|social>.xlsx"
function nomTelechargement(parrain, type) {
    return `${sanitizeFileName(parrain)}_${type}.xlsx`;
}

// Un fichier est "disponible" au téléchargement s'il existe ET n'est pas vide.
function estDisponible(fichier) {
    return Boolean(fichier) && !fichier.vide;
}

// Classe CSS du badge : présent + rempli / présent mais vide / absent.
function badgeClasse(fichier, disponible) {
    if (disponible) return "badgeOn";
    if (fichier) return "badgeEmpty";
    return "badgeOff";
}

async function telechargerUnFichier(url, nomFichier) {
    const response = await fetch(url);
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

async function telechargerPlusieurs(fichiers) {
    for (const { url, nomFichier } of fichiers) {
        await telechargerUnFichier(url, nomFichier);
    }
}

// Lit un fichier <input type="file"> et le convertit en base64 (sans le
// préfixe "data:...;base64,") pour pouvoir l'envoyer en JSON à l'API.
function fichierVersBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const resultat = String(reader.result || "");
            const base64 = resultat.split(",")[1] || "";
            resolve(base64);
        };
        reader.onerror = () => reject(new Error("Impossible de lire le fichier."));
        reader.readAsDataURL(file);
    });
}

// =========================================================
// MODALE D'IMPORT DE FICHIERS
// =========================================================

function ModaleImport({ onClose, onSuccess }) {
    const [fichiersMoltaqa, setFichiersMoltaqa] = useState([]);
    const [fichiersSocial, setFichiersSocial] = useState([]);
    const [envoiEnCours, setEnvoiEnCours] = useState(false);
    const [erreur, setErreur] = useState("");

    const totalFichiers = fichiersMoltaqa.length + fichiersSocial.length;

    async function handleSubmit(e) {
        e.preventDefault();
        setErreur("");

        if (totalFichiers === 0) {
            setErreur("Sélectionnez au moins un fichier (Moltaqa ou Social).");
            return;
        }

        setEnvoiEnCours(true);
        try {
            const moltaqa = await Promise.all(
                fichiersMoltaqa.map(async (file) => ({
                    nomFichier: file.name,
                    data: await fichierVersBase64(file),
                }))
            );
            const social = await Promise.all(
                fichiersSocial.map(async (file) => ({
                    nomFichier: file.name,
                    data: await fichierVersBase64(file),
                }))
            );

            const res = await fetch("/api/excels/add", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ moltaqa, social }),
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
                    <h2>Importer des fichiers</h2>
                    <button type="button" className="modalClose" onClick={onClose} aria-label="Fermer">
                        <X size={18} strokeWidth={2} />
                    </button>
                </div>

                <form onSubmit={handleSubmit} className="modalForm">
                    <p className="modalHint">
                        Chaque fichier est enregistré sous son propre nom (le nom du parrain
                        correspond automatiquement au nom du fichier).
                    </p>

                    <label className="fieldLabel">
                        Fichiers Moltaqa (.xlsx) — sélection multiple possible
                        <input
                            type="file"
                            accept=".xlsx,.xls"
                            multiple
                            onChange={(e) => setFichiersMoltaqa(Array.from(e.target.files || []))}
                        />
                        {fichiersMoltaqa.length > 0 && (
                            <ul className="fieldFileList">
                                {fichiersMoltaqa.map((f, i) => (
                                    <li key={`${f.name}-${i}`}>{f.name}</li>
                                ))}
                            </ul>
                        )}
                    </label>

                    <label className="fieldLabel">
                        Fichiers Social (.xlsx) — sélection multiple possible
                        <input
                            type="file"
                            accept=".xlsx,.xls"
                            multiple
                            onChange={(e) => setFichiersSocial(Array.from(e.target.files || []))}
                        />
                        {fichiersSocial.length > 0 && (
                            <ul className="fieldFileList">
                                {fichiersSocial.map((f, i) => (
                                    <li key={`${f.name}-${i}`}>{f.name}</li>
                                ))}
                            </ul>
                        )}
                    </label>

                    {erreur && <p className="modalErreur">{erreur}</p>}

                    <div className="modalActions">
                        <button type="button" className="actionBtn" onClick={onClose} disabled={envoiEnCours}>
                            Annuler
                        </button>
                        <button type="submit" className="actionBtn actionBtnPrimary" disabled={envoiEnCours}>
                            {envoiEnCours ? (
                                <Loader2 size={14} strokeWidth={2} className="spin" />
                            ) : (
                                <Plus size={14} strokeWidth={2} />
                            )}
                            {envoiEnCours
                                ? "Import..."
                                : `Importer${totalFichiers > 0 ? ` (${totalFichiers})` : ""}`}
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
          max-width: 460px;
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
          gap: 16px;
        }
        .modalHint {
          margin: 0;
          font-size: 12.5px;
          color: #8a8378;
          line-height: 1.5;
        }
        .fieldLabel {
          display: flex;
          flex-direction: column;
          gap: 6px;
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
        .fieldFileList {
          list-style: none;
          margin: 0;
          padding: 0;
          display: flex;
          flex-direction: column;
          gap: 2px;
          max-height: 110px;
          overflow-y: auto;
          background: #fbfaf8;
          border: 1px solid #e5e1d8;
          border-radius: 4px;
          padding: 6px 10px;
        }
        .fieldFileList li {
          font-size: 12px;
          font-weight: 500;
          text-transform: none;
          letter-spacing: normal;
          color: #1f3a5f;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
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

        @media (max-width: 480px) {
          .modalActions {
            flex-direction: column-reverse;
          }
          .modalActions .actionBtn {
            width: 100%;
          }
        }
      `}</style>
        </div>
    );
}

// =========================================================
// PAGE
// =========================================================

export default function ExcelsPage({ parrains }) {
    const router = useRouter();
    const [search, setSearch] = useState("");
    const [modaleOuverte, setModaleOuverte] = useState(false);
    const [suppressionEnCours, setSuppressionEnCours] = useState(null);

    const parrainsFiltres = useMemo(() => {
        const q = normaliser(search);
        if (!q) return parrains;
        return parrains.filter((p) => normaliser(p.nom).includes(q));
    }, [parrains, search]);

    const totalMoltaqa = parrains.filter((p) => estDisponible(p.moltaqa)).length;
    const totalSocial = parrains.filter((p) => estDisponible(p.social)).length;
    const totalComplet = parrains.filter(
        (p) => estDisponible(p.moltaqa) && estDisponible(p.social)
    ).length;

    // Redemande les données au serveur (getServerSideProps) sans rechargement complet.
    function rafraichir() {
        router.replace(router.asPath, undefined, { scroll: false });
    }

    async function handleSupprimer(nomParrain) {
        const confirme = window.confirm(
            `Supprimer définitivement "${nomParrain}" (fichiers Moltaqa et Social) ?`
        );
        if (!confirme) return;

        setSuppressionEnCours(nomParrain);
        try {
            const res = await fetch("/api/excels/delete", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ nom: nomParrain }),
            });
            const data = await res.json().catch(() => ({}));
            if (!res.ok) {
                throw new Error(data.error || "Erreur lors de la suppression.");
            }
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
                <header className="letterhead">
                    <div className="letterheadBar" />
                    <div className="letterheadText">
                        <span className="eyebrow">Dossiers Excel</span>
                        <h1>Fichiers par parrain</h1>
                        <p>Retrouvez et téléchargez les fichiers Moltaqa et Social de chaque parrain.</p>
                    </div>
                </header>

                <section className="statsRow">
                    <div className="statCard">
                        <span className="statValue">{parrains.length}</span>
                        <span className="statLabel">Parrains</span>
                    </div>
                    <div className="statCard">
                        <span className="statValue">{totalMoltaqa}</span>
                        <span className="statLabel">Moltaqa</span>
                    </div>
                    <div className="statCard">
                        <span className="statValue">{totalSocial}</span>
                        <span className="statLabel">Social</span>
                    </div>
                    <div className="statCard">
                        <span className="statValue">{totalComplet}</span>
                        <span className="statLabel">Complets</span>
                    </div>
                </section>

                <section className="searchSection">
                    <div className="searchRow">
                        <div className="searchBox">
                            <Search size={16} strokeWidth={2} className="searchIcon" />
                            <input
                                type="text"
                                placeholder="Rechercher un parrain..."
                                value={search}
                                onChange={(e) => setSearch(e.target.value)}
                            />
                        </div>
                        <button
                            type="button"
                            className="actionBtn actionBtnPrimary addBtn"
                            onClick={() => setModaleOuverte(true)}
                        >
                            <Plus size={14} strokeWidth={2} />
                            Importer des fichiers
                        </button>
                    </div>
                </section>

                <section className="listSection">
                    {parrainsFiltres.length === 0 ? (
                        <div className="empty">
                            <FolderOpen size={30} strokeWidth={1.6} />
                            <p>
                                {parrains.length === 0
                                    ? "Aucun fichier trouvé dans les dossiers Moltaqa ou Social."
                                    : "Aucun parrain ne correspond à votre recherche."}
                            </p>
                        </div>
                    ) : (
                        <div className="parrainList">
                            {parrainsFiltres.map((p) => {
                                const moltaqaOk = estDisponible(p.moltaqa);
                                const socialOk = estDisponible(p.social);
                                const nbDisponibles = (moltaqaOk ? 1 : 0) + (socialOk ? 1 : 0);
                                const enSuppression = suppressionEnCours === p.nom;

                                const fichiersATelecharger = [
                                    moltaqaOk && {
                                        url: p.moltaqa.url,
                                        nomFichier: nomTelechargement(p.nom, "moltaqa"),
                                    },
                                    socialOk && {
                                        url: p.social.url,
                                        nomFichier: nomTelechargement(p.nom, "social"),
                                    },
                                ].filter(Boolean);

                                let libelleCombine = "Les deux";
                                if (nbDisponibles === 0) libelleCombine = "Indisponible";
                                else if (nbDisponibles === 1) libelleCombine = "Télécharger le disponible";

                                return (
                                    <div key={p.nom} className="parrainCard">
                                        <div className="parrainInfo">
                                            <div className="parrainIcon">
                                                <FileSpreadsheet size={18} strokeWidth={1.8} />
                                            </div>
                                            <div className="parrainText">
                                                <span className="parrainNom">{p.nom}</span>
                                                <div className="parrainBadges">
                                                    <span
                                                        className={`badge ${badgeClasse(p.moltaqa, moltaqaOk)}`}
                                                    >
                                                        Moltaqa{p.moltaqa && !moltaqaOk ? " (vide)" : ""}
                                                    </span>
                                                    <span className={`badge ${badgeClasse(p.social, socialOk)}`}>
                                                        Social{p.social && !socialOk ? " (vide)" : ""}
                                                    </span>
                                                </div>
                                            </div>
                                        </div>

                                        <div className="parrainActions">
                                            <button
                                                type="button"
                                                className={`actionBtn ${!moltaqaOk ? "actionBtnDisabled" : ""}`}
                                                disabled={!moltaqaOk}
                                                onClick={() =>
                                                    moltaqaOk &&
                                                    telechargerUnFichier(p.moltaqa.url, nomTelechargement(p.nom, "moltaqa"))
                                                }
                                            >
                                                <Download size={14} strokeWidth={2} />
                                                Moltaqa
                                            </button>

                                            <button
                                                type="button"
                                                className={`actionBtn ${!socialOk ? "actionBtnDisabled" : ""}`}
                                                disabled={!socialOk}
                                                onClick={() =>
                                                    socialOk &&
                                                    telechargerUnFichier(p.social.url, nomTelechargement(p.nom, "social"))
                                                }
                                            >
                                                <Download size={14} strokeWidth={2} />
                                                Social
                                            </button>

                                            <button
                                                type="button"
                                                className="actionBtn actionBtnPrimary"
                                                disabled={nbDisponibles === 0}
                                                onClick={() => telechargerPlusieurs(fichiersATelecharger)}
                                            >
                                                <Download size={14} strokeWidth={2} />
                                                {libelleCombine}
                                            </button>

                                            <button
                                                type="button"
                                                className="actionBtn actionBtnDanger"
                                                disabled={enSuppression}
                                                onClick={() => handleSupprimer(p.nom)}
                                                aria-label={`Supprimer ${p.nom}`}
                                            >
                                                {enSuppression ? (
                                                    <Loader2 size={14} strokeWidth={2} className="spin" />
                                                ) : (
                                                    <Trash2 size={14} strokeWidth={2} />
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
                <ModaleImport
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
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          color: #1e2124;
        }
        .sheet {
          width: 100%;
          max-width: 900px;
          background: #ffffff;
          border: 1px solid #e0ddd4;
          border-radius: 4px;
        }

        /* ============ HEADER ============ */

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

        /* ============ STATS ============ */

        .statsRow {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
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

        /* ============ SEARCH ============ */

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
        .searchIcon {
          position: absolute;
          left: 12px;
          color: #9a9488;
          pointer-events: none;
        }
        .searchBox input {
          width: 100%;
          background: #fbfaf8;
          border: 1px solid #d7d2c6;
          color: #1e2124;
          border-radius: 4px;
          padding: 10px 12px 10px 36px;
          font-size: 14px;
          box-sizing: border-box;
        }
        .searchBox input:focus {
          outline: none;
          border-color: #1f3a5f;
          box-shadow: 0 0 0 3px rgba(31, 58, 95, 0.12);
        }
        .addBtn {
          white-space: nowrap;
          flex-shrink: 0;
        }

        /* ============ LIST ============ */

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
        .badgeOff {
          background: #f1efe9;
          color: #b3ac9f;
        }
        .badgeEmpty {
          background: #fbf1e4;
          color: #a3742f;
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
        .actionBtn:hover:not(.actionBtnDisabled):not(:disabled) {
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
        .actionBtnPrimary:disabled {
          background: #d7d2c6;
          border-color: #d7d2c6;
          color: #9a9488;
          cursor: not-allowed;
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

        /* ============ RESPONSIVE ============ */

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
            grid-template-columns: 1fr 1fr;
          }
          .parrainActions .actionBtnPrimary {
            grid-column: 1 / -1;
          }
          .parrainActions .actionBtnDanger {
            grid-column: 1 / -1;
          }
          .actionBtn {
            flex: 1;
            width: 100%;
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
