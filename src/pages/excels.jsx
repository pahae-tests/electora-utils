import { useMemo, useState } from "react";
import fs from "fs";
import path from "path";
import { FileSpreadsheet, Download, Search, FolderOpen } from "lucide-react";

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

      if (!parrainsMap.has(nomParrain)) {
        parrainsMap.set(nomParrain, { nom: nomParrain });
      }

      parrainsMap.get(nomParrain)[key] = `/${dir}/${fichier}`;
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

// Déclenche plusieurs téléchargements successifs (avec un léger délai pour
// éviter que le navigateur ne bloque le second en tant que pop-up), en
// forçant le nom de chaque fichier téléchargé.
function telechargerPlusieurs(fichiers) {
  fichiers.forEach(({ url, nomFichier }, index) => {
    setTimeout(() => {
      const a = document.createElement("a");
      a.href = url;
      a.download = nomFichier;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }, index * 350);
  });
}

// =========================================================
// PAGE
// =========================================================

export default function ExcelsPage({ parrains }) {
  const [search, setSearch] = useState("");

  const parrainsFiltres = useMemo(() => {
    const q = normaliser(search);
    if (!q) return parrains;
    return parrains.filter((p) => normaliser(p.nom).includes(q));
  }, [parrains, search]);

  const totalMoltaqa = parrains.filter((p) => p.moltaqa).length;
  const totalSocial = parrains.filter((p) => p.social).length;
  const totalComplet = parrains.filter((p) => p.moltaqa && p.social).length;

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
          <div className="searchBox">
            <Search size={16} strokeWidth={2} className="searchIcon" />
            <input
              type="text"
              placeholder="Rechercher un parrain..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
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
              {parrainsFiltres.map((p) => (
                <div key={p.nom} className="parrainCard">
                  <div className="parrainInfo">
                    <div className="parrainIcon">
                      <FileSpreadsheet size={18} strokeWidth={1.8} />
                    </div>
                    <div className="parrainText">
                      <span className="parrainNom">{p.nom}</span>
                      <div className="parrainBadges">
                        <span className={`badge ${p.moltaqa ? "badgeOn" : "badgeOff"}`}>
                          Moltaqa
                        </span>
                        <span className={`badge ${p.social ? "badgeOn" : "badgeOff"}`}>
                          Social
                        </span>
                      </div>
                    </div>
                  </div>

                  <div className="parrainActions">
                    <a
                      className={`actionBtn ${!p.moltaqa ? "actionBtnDisabled" : ""}`}
                      href={p.moltaqa || undefined}
                      download={p.moltaqa ? nomTelechargement(p.nom, "moltaqa") : undefined}
                      aria-disabled={!p.moltaqa}
                      onClick={(e) => !p.moltaqa && e.preventDefault()}
                    >
                      <Download size={14} strokeWidth={2} />
                      Moltaqa
                    </a>

                    <a
                      className={`actionBtn ${!p.social ? "actionBtnDisabled" : ""}`}
                      href={p.social || undefined}
                      download={p.social ? nomTelechargement(p.nom, "social") : undefined}
                      aria-disabled={!p.social}
                      onClick={(e) => !p.social && e.preventDefault()}
                    >
                      <Download size={14} strokeWidth={2} />
                      Social
                    </a>

                    <button
                      type="button"
                      className="actionBtn actionBtnPrimary"
                      disabled={!p.moltaqa || !p.social}
                      onClick={() =>
                        telechargerPlusieurs([
                          { url: p.moltaqa, nomFichier: nomTelechargement(p.nom, "moltaqa") },
                          { url: p.social, nomFichier: nomTelechargement(p.nom, "social") },
                        ])
                      }
                    >
                      <Download size={14} strokeWidth={2} />
                      Les deux
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </section>
      </div>

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
        .searchBox {
          position: relative;
          display: flex;
          align-items: center;
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

        .parrainActions {
          display: flex;
          gap: 8px;
          flex-shrink: 0;
        }

        .actionBtn {
          display: inline-flex;
          align-items: center;
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
          background: #16293f;
        }
        .actionBtnPrimary:disabled {
          background: #d7d2c6;
          border-color: #d7d2c6;
          color: #9a9488;
          cursor: not-allowed;
        }

        @media (max-width: 640px) {
          .letterheadText,
          .searchSection,
          .listSection {
            padding-left: 22px;
            padding-right: 22px;
          }
          .statsRow {
            grid-template-columns: repeat(2, 1fr);
          }
          .parrainCard {
            flex-direction: column;
            align-items: flex-start;
          }
          .parrainActions {
            width: 100%;
          }
          .actionBtn {
            flex: 1;
            justify-content: center;
          }
        }
      `}</style>
    </div>
  );
}