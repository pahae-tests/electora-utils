import { useState, useRef } from "react";
import * as XLSX from "xlsx";

// Colonnes à conserver depuis le fichier importé (comparaison insensible
// à la casse / aux espaces pour tolérer de légères variations d'en-tête).
const COLONNES_CONSERVEES = ["N° Électeur", "Nom", "Prénom", "Téléphone", "Bureau de vote", "Parrain"];

const SERVICE_OPTIONS = [
  { value: "social", label: "Social" },
  { value: "molta9a", label: "Motla9a" },
];

function normaliser(txt) {
  return String(txt ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function sanitizeFileName(name) {
  const cleaned = String(name || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || "export";
}

export default function Home() {
  const fileInputRef = useRef(null);

  const [fileName, setFileName] = useState("");
  const [rawHeaders, setRawHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [serviceSocial, setServiceSocial] = useState(SERVICE_OPTIONS[0].value);
  const [resultRows, setResultRows] = useState(null);
  const [error, setError] = useState("");

  const handleFileChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setError("");
    setResultRows(null);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        if (!rows.length) {
          setError("Le fichier importé est vide.");
          return;
        }

        const [headerRow, ...dataRows] = rows;
        setRawHeaders(headerRow);
        setRawRows(dataRows);
      } catch (err) {
        setError("Impossible de lire ce fichier. Vérifiez qu'il s'agit bien d'un fichier Excel valide.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleGenerer = () => {
    if (!rawRows.length) {
      setError("Importez d'abord un fichier Excel.");
      return;
    }
    setError("");

    // Repère l'index de chaque colonne à conserver dans l'en-tête importé.
    const indexParColonne = COLONNES_CONSERVEES.map((col) =>
      rawHeaders.findIndex((h) => normaliser(h) === normaliser(col))
    );

    const colonnesManquantes = COLONNES_CONSERVEES.filter((_, i) => indexParColonne[i] === -1);
    if (colonnesManquantes.length) {
      setError(`Colonnes introuvables dans le fichier : ${colonnesManquantes.join(", ")}`);
      return;
    }

    const serviceLabel = SERVICE_OPTIONS.find((o) => o.value === serviceSocial)?.label ?? "";

    const lignes = rawRows
      // Ignore uniquement les lignes totalement vides.
      .filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""))
      .map((row) => {
        const nonVides = row.filter((cell) => String(cell ?? "").trim() !== "");
        // Ligne récapitulative (ex. "Total : 36 électeur(s)") : une seule cellule
        // renseignée sur toute la ligne. On la conserve telle quelle, sans lui
        // affecter de Service social.
        const estLigneTotal = nonVides.length === 1 && /^total\s*:/i.test(String(nonVides[0]).trim());

        const obj = {};
        if (estLigneTotal) {
          COLONNES_CONSERVEES.forEach((col, i) => {
            obj[col] = "";
          });
          obj[COLONNES_CONSERVEES[0]] = nonVides[0];
          obj["Service social"] = "";
        } else {
          COLONNES_CONSERVEES.forEach((col, i) => {
            obj[col] = row[indexParColonne[i]] ?? "";
          });
          obj["Service social"] = serviceLabel;
        }
        obj.__isTotal = estLigneTotal;
        return obj;
      });

    // Suppression des doublons (mêmes valeurs sur les 5 colonnes conservées),
    // la ligne récapitulative de total est toujours conservée.
    const vus = new Set();
    const lignesUniques = [];
    for (const ligne of lignes) {
      if (ligne.__isTotal) {
        lignesUniques.push(ligne);
        continue;
      }
      const cle = COLONNES_CONSERVEES.map((col) => normaliser(ligne[col])).join("|");
      if (!vus.has(cle)) {
        vus.add(cle);
        lignesUniques.push(ligne);
      }
    }

    setResultRows(lignesUniques);
  };

  const handleExporter = () => {
    if (!resultRows || !resultRows.length) return;

    const parrain = resultRows.find((l) => String(l["Parrain"]).trim() !== "")?.["Parrain"] ?? "export";
    const nomFichier = `${sanitizeFileName(parrain)}.xlsx`;

    const lignesExport = resultRows.map(({ __isTotal, ...ligne }) => ligne);

    const worksheet = XLSX.utils.json_to_sheet(lignesExport, {
      header: [...COLONNES_CONSERVEES, "Service social"],
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Électeurs");
    XLSX.writeFile(workbook, nomFichier);
  };

  const resetFichier = () => {
    setFileName("");
    setRawHeaders([]);
    setRawRows([]);
    setResultRows(null);
    setError("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  return (
    <div className="page">
      <div className="sheet">
        <header className="letterhead">
          <div className="letterheadBar" />
          <div className="letterheadText">
            <span className="eyebrow">Registre électoral</span>
            <h1>Répartition des électeurs</h1>
            <p>Importez la liste, affectez un service, générez le registre filtré.</p>
          </div>
        </header>

        <section className="formGrid">
          <div className="field">
            <label htmlFor="fichier">Fichier Excel</label>
            <div className="fileRow">
              <input
                id="fichier"
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileChange}
              />
            </div>
            {fileName && (
              <div className="fileInfo">
                <span>{fileName}</span>
                <button type="button" className="linkBtn" onClick={resetFichier}>
                  Retirer
                </button>
              </div>
            )}
          </div>

          <div className="field">
            <label htmlFor="service">Service social</label>
            <select
              id="service"
              value={serviceSocial}
              onChange={(e) => setServiceSocial(e.target.value)}
            >
              {SERVICE_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
          </div>
        </section>

        {error && <div className="error">{error}</div>}

        <button type="button" className="primaryBtn" onClick={handleGenerer} disabled={!rawRows.length}>
          Générer
        </button>

        {resultRows && (
          <section className="results">
            <div className="resultsHeader">
              <h2>Registre généré</h2>
              <button type="button" className="secondaryBtn" onClick={handleExporter} disabled={!resultRows.length}>
                Exporter le fichier
              </button>
            </div>

            <div className="tableWrap">
              <table>
                <thead>
                  <tr>
                    {COLONNES_CONSERVEES.map((col) => (
                      <th key={col}>{col}</th>
                    ))}
                    <th>Service social</th>
                  </tr>
                </thead>
                <tbody>
                  {resultRows.map((ligne, i) => (
                    <tr key={i}>
                      {COLONNES_CONSERVEES.map((col) => (
                        <td key={col}>{ligne[col]}</td>
                      ))}
                      <td>
                        {ligne["Service social"] && <span className="tag">{ligne["Service social"]}</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="totalRow">
              <span>Total électeurs</span>
              <strong>{resultRows.filter((l) => !l.__isTotal).length}</strong>
            </div>
          </section>
        )}
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
          max-width: 760px;
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
        .formGrid {
          padding: 28px 36px 8px;
          display: grid;
          grid-template-columns: 1fr 220px;
          gap: 24px;
        }
        .field {
          display: flex;
          flex-direction: column;
          gap: 8px;
        }
        label {
          font-size: 12px;
          font-weight: 600;
          color: #4a4740;
          letter-spacing: 0.01em;
        }
        input[type="file"] {
          font-size: 13px;
          color: #4a4740;
        }
        select {
          background: #fbfaf8;
          border: 1px solid #d7d2c6;
          color: #1e2124;
          border-radius: 4px;
          padding: 9px 10px;
          font-size: 14px;
        }
        select:focus,
        input:focus {
          outline: none;
          border-color: #1f3a5f;
          box-shadow: 0 0 0 3px rgba(31, 58, 95, 0.12);
        }
        .fileInfo {
          display: flex;
          align-items: center;
          gap: 12px;
          font-size: 12px;
          color: #8a8378;
        }
        .linkBtn {
          background: none;
          border: none;
          color: #1f3a5f;
          cursor: pointer;
          font-size: 12px;
          padding: 0;
          text-decoration: underline;
        }
        .primaryBtn {
          margin: 20px 36px 4px;
          background: #1f3a5f;
          color: #ffffff;
          border: none;
          border-radius: 4px;
          padding: 11px 22px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
        }
        .primaryBtn:hover:not(:disabled) {
          background: #16293f;
        }
        .primaryBtn:disabled {
          background: #d7d2c6;
          color: #9a9488;
          cursor: not-allowed;
        }
        .secondaryBtn {
          background: transparent;
          border: 1px solid #1f3a5f;
          color: #1f3a5f;
          border-radius: 4px;
          padding: 8px 16px;
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
        }
        .secondaryBtn:hover:not(:disabled) {
          background: #eef2f6;
        }
        .secondaryBtn:disabled {
          opacity: 0.4;
          cursor: not-allowed;
        }
        .error {
          margin: 20px 36px 4px;
          background: #fbeeee;
          border: 1px solid #e3b8b8;
          color: #8a2d2d;
          border-radius: 4px;
          padding: 10px 14px;
          font-size: 13px;
        }
        .results {
          margin: 32px 36px 36px;
          border-top: 1px solid #e0ddd4;
          padding-top: 24px;
        }
        .resultsHeader {
          display: flex;
          align-items: center;
          justify-content: space-between;
          margin-bottom: 14px;
        }
        .resultsHeader h2 {
          font-family: Georgia, "Times New Roman", serif;
          font-size: 17px;
          margin: 0;
          font-weight: 600;
          color: #16191c;
        }
        .tableWrap {
          overflow-x: auto;
          border: 1px solid #e0ddd4;
          border-radius: 4px;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          font-size: 13px;
        }
        th {
          text-align: left;
          padding: 10px 14px;
          background: #fbfaf8;
          color: #6b6459;
          font-weight: 600;
          font-size: 12px;
          border-bottom: 2px solid #1f3a5f;
          white-space: nowrap;
        }
        td {
          text-align: left;
          padding: 9px 14px;
          border-bottom: 1px solid #ede9e0;
          white-space: nowrap;
          color: #1e2124;
        }
        tbody tr:nth-child(even) {
          background: #fbfaf8;
        }
        tbody tr:last-child td {
          border-bottom: none;
        }
        .tag {
          display: inline-block;
          background: #eef2f6;
          color: #1f3a5f;
          border-radius: 3px;
          padding: 2px 8px;
          font-size: 12px;
          font-weight: 600;
        }
        .totalRow {
          margin-top: 16px;
          padding-top: 12px;
          border-top: 2px solid #1f3a5f;
          display: flex;
          justify-content: space-between;
          align-items: baseline;
          font-size: 14px;
          color: #4a4740;
        }
        .totalRow strong {
          font-family: Georgia, "Times New Roman", serif;
          font-size: 20px;
          color: #16191c;
        }
        @media (max-width: 560px) {
          .formGrid {
            grid-template-columns: 1fr;
          }
          .letterheadText,
          .formGrid,
          .results {
            padding-left: 22px;
            padding-right: 22px;
          }
          .primaryBtn {
            margin-left: 22px;
            margin-right: 22px;
          }
          .error {
            margin-left: 22px;
            margin-right: 22px;
          }
        }
      `}</style>
    </div>
  );
}