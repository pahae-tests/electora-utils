import { useState, useRef, useMemo } from "react";
import * as XLSX from "xlsx";

export default function Check() {
  const fileInputRef = useRef(null);

  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [search, setSearch] = useState("");
  const [selectedColumn, setSelectedColumn] = useState("__ALL__");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  // =========================
  // NORMALISATION
  // =========================

  const normaliser = (value) => {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  };

  // =========================
  // IMPORT EXCEL
  // =========================

  const handleFileChange = (e) => {
    const file = e.target.files?.[0];

    if (!file) return;

    setError("");
    setLoading(true);
    setFileName(file.name);
    setHeaders([]);
    setRows([]);
    setSearch("");
    setSelectedColumn("__ALL__");

    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);

        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
        });

        if (!workbook.SheetNames.length) {
          throw new Error("Aucune feuille trouvée.");
        }

        // Première feuille
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Conversion en tableau
        const json = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
          raw: false,
        });

        if (!json.length) {
          throw new Error("Le fichier Excel est vide.");
        }

        // Première ligne = en-têtes
        const headerRow = json[0];

        // Nettoyage des en-têtes
        const cleanHeaders = headerRow.map((header, index) => {
          const value = String(header ?? "").trim();

          return value || `Colonne ${index + 1}`;
        });

        // Évite les noms de colonnes identiques
        const uniqueHeaders = [];
        const headerCount = {};

        cleanHeaders.forEach((header) => {
          if (!headerCount[header]) {
            headerCount[header] = 1;
            uniqueHeaders.push(header);
          } else {
            headerCount[header]++;
            uniqueHeaders.push(
              `${header} (${headerCount[header]})`
            );
          }
        });

        // Transformation en objets
        const dataRows = json
          .slice(1)
          .filter((row) =>
            row.some(
              (cell) => String(cell ?? "").trim() !== ""
            )
          )
          .map((row) => {
            const obj = {};

            uniqueHeaders.forEach((header, index) => {
              obj[header] = row[index] ?? "";
            });

            return obj;
          });

        setHeaders(uniqueHeaders);
        setRows(dataRows);

        if (!dataRows.length) {
          setError("Le fichier contient des en-têtes mais aucune donnée.");
        }
      } catch (err) {
        console.error(err);

        setError(
          err.message ||
            "Impossible de lire ce fichier Excel. Vérifiez qu'il est valide."
        );
      } finally {
        setLoading(false);
      }
    };

    reader.onerror = () => {
      setError("Impossible de lire le fichier.");
      setLoading(false);
    };

    reader.readAsArrayBuffer(file);
  };

  // =========================
  // RECHERCHE
  // =========================

  const filteredRows = useMemo(() => {
    const query = normaliser(search);

    if (!query) {
      return rows;
    }

    return rows.filter((row) => {
      // Recherche dans toutes les colonnes
      if (selectedColumn === "__ALL__") {
        return headers.some((header) =>
          normaliser(row[header]).includes(query)
        );
      }

      // Recherche dans une seule colonne
      return normaliser(row[selectedColumn]).includes(query);
    });
  }, [rows, headers, search, selectedColumn]);

  // =========================
  // RESET
  // =========================

  const resetFile = () => {
    setFileName("");
    setHeaders([]);
    setRows([]);
    setSearch("");
    setSelectedColumn("__ALL__");
    setError("");

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const resetSearch = () => {
    setSearch("");
  };

  // =========================
  // AFFICHAGE
  // =========================

  return (
    <div className="page">
      <div className="sheet">

        {/* ================= HEADER ================= */}

        <header className="letterhead">
          <div className="letterheadBar" />

          <div className="letterheadText">
            <span className="eyebrow">
              Vérification électorale
            </span>

            <h1>Recherche dans le registre</h1>

            <p>
              Importez un fichier Excel et recherchez rapidement
              un électeur dans n'importe quelle colonne.
            </p>
          </div>
        </header>

        {/* ================= IMPORT ================= */}

        <section className="importSection">

          <div className="field">
            <label htmlFor="fichier">
              Fichier Excel
            </label>

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

                <button
                  type="button"
                  className="linkBtn"
                  onClick={resetFile}
                >
                  Retirer
                </button>
              </div>
            )}
          </div>

        </section>

        {/* ================= ERROR ================= */}

        {error && (
          <div className="error">
            {error}
          </div>
        )}

        {/* ================= LOADING ================= */}

        {loading && (
          <div className="loading">
            Lecture du fichier...
          </div>
        )}

        {/* ================= SEARCH ================= */}

        {headers.length > 0 && rows.length > 0 && (
          <section className="searchSection">

            <div className="searchHeader">
              <div>
                <span className="sectionLabel">
                  Recherche
                </span>

                <h2>
                  Trouver un électeur
                </h2>
              </div>

              <div className="stats">
                <span>
                  {rows.length.toLocaleString("fr-FR")}
                </span>

                <small>
                  électeurs
                </small>
              </div>
            </div>

            <div className="searchGrid">

              {/* COLONNE */}

              <div className="field">
                <label htmlFor="column">
                  Rechercher dans
                </label>

                <select
                  id="column"
                  value={selectedColumn}
                  onChange={(e) =>
                    setSelectedColumn(e.target.value)
                  }
                >
                  <option value="__ALL__">
                    Toutes les colonnes
                  </option>

                  {headers.map((header, index) => (
                    <option
                      key={`${header}-${index}`}
                      value={header}
                    >
                      {header}
                    </option>
                  ))}
                </select>
              </div>

              {/* RECHERCHE */}

              <div className="field searchField">
                <label htmlFor="search">
                  Terme de recherche
                </label>

                <div className="searchInputWrapper">

                  <span className="searchIcon">
                    ⌕
                  </span>

                  <input
                    id="search"
                    type="text"
                    value={search}
                    onChange={(e) =>
                      setSearch(e.target.value)
                    }
                    placeholder={
                      selectedColumn === "__ALL__"
                        ? "Nom, CIN, téléphone, bureau..."
                        : `Rechercher dans « ${selectedColumn} »`
                    }
                  />

                  {search && (
                    <button
                      type="button"
                      className="clearSearch"
                      onClick={resetSearch}
                      aria-label="Effacer"
                    >
                      ×
                    </button>
                  )}

                </div>
              </div>

            </div>

            {/* ================= SEARCH INFO ================= */}

            <div className="searchInfo">

              <span>
                {search ? (
                  <>
                    Résultats pour{" "}
                    <strong>« {search} »</strong>
                  </>
                ) : (
                  "Toutes les lignes sont affichées"
                )}
              </span>

              <strong className="resultCount">
                {filteredRows.length.toLocaleString("fr-FR")}
                {" "}
                résultat
                {filteredRows.length !== 1 ? "s" : ""}
              </strong>

            </div>

          </section>
        )}

        {/* ================= RESULTS ================= */}

        {headers.length > 0 && rows.length > 0 && (
          <section className="results">

            <div className="resultsHeader">

              <div>
                <h2>
                  Registre
                </h2>

                <p>
                  {filteredRows.length === rows.length
                    ? `${rows.length.toLocaleString("fr-FR")} lignes`
                    : `${filteredRows.length.toLocaleString(
                        "fr-FR"
                      )} sur ${rows.length.toLocaleString(
                        "fr-FR"
                      )} lignes`}
                </p>
              </div>

              {search && (
                <button
                  type="button"
                  className="secondaryBtn"
                  onClick={resetSearch}
                >
                  Réinitialiser
                </button>
              )}

            </div>

            <div className="tableWrap">

              <table>

                <thead>
                  <tr>
                    <th className="numberColumn">
                      #
                    </th>

                    {headers.map((header, index) => (
                      <th key={`${header}-${index}`}>
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>

                <tbody>

                  {filteredRows.length > 0 ? (
                    filteredRows.map((row, rowIndex) => (
                      <tr key={rowIndex}>

                        <td className="rowNumber">
                          {rows.indexOf(row) + 1}
                        </td>

                        {headers.map((header, index) => (
                          <td
                            key={`${header}-${index}`}
                          >
                            {String(row[header] ?? "")}
                          </td>
                        ))}

                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td
                        colSpan={headers.length + 1}
                        className="noResults"
                      >
                        <div className="noResultsTitle">
                          Aucun résultat
                        </div>

                        <div className="noResultsText">
                          Aucun électeur ne correspond à votre
                          recherche.
                        </div>
                      </td>
                    </tr>
                  )}

                </tbody>

              </table>

            </div>

            {/* ================= TOTAL ================= */}

            <div className="totalRow">

              <span>
                Électeurs affichés
              </span>

              <strong>
                {filteredRows.length.toLocaleString("fr-FR")}
              </strong>

            </div>

          </section>
        )}

        {/* ================= EMPTY STATE ================= */}

        {!headers.length && !loading && (
          <div className="emptyState">

            <div className="emptyIcon">
              XLS
            </div>

            <h2>
              Importez votre fichier
            </h2>

            <p>
              Sélectionnez un fichier Excel pour commencer
              la recherche.
            </p>

          </div>
        )}

      </div>

      {/* ================= STYLE ================= */}

      <style jsx>{`

        .page {
          min-height: 100vh;
          background: #f4f2ee;
          padding: 56px 20px;
          display: flex;
          justify-content: center;
          font-family:
            -apple-system,
            BlinkMacSystemFont,
            "Segoe UI",
            Roboto,
            sans-serif;
          color: #1e2124;
        }

        .sheet {
          width: 100%;
          max-width: 1200px;
          background: #ffffff;
          border: 1px solid #e0ddd4;
          border-radius: 4px;
          overflow: hidden;
        }

        /* ================= HEADER ================= */

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
          font-family:
            Georgia,
            "Times New Roman",
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

        /* ================= IMPORT ================= */

        .importSection {
          padding: 28px 36px 10px;
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

        input[type="file"]::file-selector-button {
          background: #fbfaf8;
          border: 1px solid #d7d2c6;
          color: #1f3a5f;
          border-radius: 4px;
          padding: 9px 14px;
          margin-right: 10px;
          cursor: pointer;
          font-weight: 600;
        }

        .fileInfo {
          display: flex;
          align-items: center;
          gap: 12px;
          margin-top: 8px;
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

        /* ================= ERROR ================= */

        .error {
          margin: 20px 36px 4px;
          background: #fbeeee;
          border: 1px solid #e3b8b8;
          color: #8a2d2d;
          border-radius: 4px;
          padding: 10px 14px;
          font-size: 13px;
        }

        .loading {
          margin: 20px 36px;
          color: #6b6459;
          font-size: 13px;
        }

        /* ================= SEARCH ================= */

        .searchSection {
          margin: 28px 36px 0;
          padding: 24px 0;
          border-top: 1px solid #e0ddd4;
          border-bottom: 1px solid #e0ddd4;
        }

        .searchHeader {
          display: flex;
          justify-content: space-between;
          align-items: flex-start;
          margin-bottom: 18px;
        }

        .sectionLabel {
          color: #96723a;
          font-size: 11px;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: 0.06em;
        }

        .searchHeader h2 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 18px;
          margin: 4px 0 0;
          color: #16191c;
        }

        .stats {
          display: flex;
          align-items: baseline;
          gap: 6px;
          color: #6b6459;
        }

        .stats span {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          color: #1f3a5f;
          font-size: 22px;
          font-weight: 600;
        }

        .stats small {
          font-size: 12px;
        }

        .searchGrid {
          display: grid;
          grid-template-columns: 260px 1fr;
          gap: 18px;
        }

        select,
        input[type="text"] {
          width: 100%;
          box-sizing: border-box;
          background: #fbfaf8;
          border: 1px solid #d7d2c6;
          color: #1e2124;
          border-radius: 4px;
          padding: 10px 11px;
          font-size: 14px;
        }

        select:focus,
        input:focus {
          outline: none;
          border-color: #1f3a5f;
          box-shadow:
            0 0 0 3px
            rgba(31, 58, 95, 0.12);
        }

        .searchInputWrapper {
          position: relative;
        }

        .searchInputWrapper input {
          padding-left: 36px;
          padding-right: 38px;
        }

        .searchIcon {
          position: absolute;
          left: 12px;
          top: 50%;
          transform: translateY(-50%);
          color: #8a8378;
          font-size: 20px;
          pointer-events: none;
        }

        .clearSearch {
          position: absolute;
          right: 9px;
          top: 50%;
          transform: translateY(-50%);
          width: 25px;
          height: 25px;
          border: none;
          background: transparent;
          color: #6b6459;
          font-size: 20px;
          cursor: pointer;
          line-height: 20px;
        }

        .clearSearch:hover {
          color: #1f3a5f;
        }

        .searchInfo {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-top: 14px;
          font-size: 12px;
          color: #8a8378;
        }

        .searchInfo strong {
          color: #4a4740;
        }

        .resultCount {
          color: #1f3a5f !important;
        }

        /* ================= RESULTS ================= */

        .results {
          margin: 28px 36px 36px;
        }

        .resultsHeader {
          display: flex;
          align-items: center;
          justify-content: space-between;
          margin-bottom: 14px;
        }

        .resultsHeader h2 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 17px;
          margin: 0;
          font-weight: 600;
          color: #16191c;
        }

        .resultsHeader p {
          margin: 4px 0 0;
          color: #8a8378;
          font-size: 12px;
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

        .secondaryBtn:hover {
          background: #eef2f6;
        }

        /* ================= TABLE ================= */

        .tableWrap {
          overflow: auto;
          max-height: 620px;
          border: 1px solid #e0ddd4;
          border-radius: 4px;
        }

        table {
          width: 100%;
          border-collapse: collapse;
          font-size: 13px;
        }

        th {
          position: sticky;
          top: 0;
          z-index: 2;
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

        tbody tr:hover {
          background: #f1f3f5;
        }

        tbody tr:last-child td {
          border-bottom: none;
        }

        .numberColumn {
          width: 50px;
          text-align: center;
        }

        .rowNumber {
          color: #9a9488;
          text-align: center;
          font-size: 11px;
        }

        /* ================= NO RESULTS ================= */

        .noResults {
          text-align: center;
          padding: 45px 20px !important;
          white-space: normal;
        }

        .noResultsTitle {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 17px;
          color: #16191c;
          margin-bottom: 6px;
        }

        .noResultsText {
          color: #8a8378;
          font-size: 13px;
        }

        /* ================= TOTAL ================= */

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
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 20px;
          color: #16191c;
        }

        /* ================= EMPTY ================= */

        .emptyState {
          margin: 35px 36px 40px;
          padding: 45px 20px;
          border: 1px dashed #d7d2c6;
          text-align: center;
          background: #fbfaf8;
        }

        .emptyIcon {
          width: 48px;
          height: 48px;
          margin: 0 auto 14px;
          display: flex;
          align-items: center;
          justify-content: center;
          border: 1px solid #d7d2c6;
          border-radius: 4px;
          color: #1f3a5f;
          font-size: 11px;
          font-weight: 700;
          letter-spacing: 0.05em;
        }

        .emptyState h2 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 18px;
          font-weight: 600;
          margin: 0 0 7px;
          color: #16191c;
        }

        .emptyState p {
          margin: 0;
          color: #8a8378;
          font-size: 13px;
        }

        /* ================= RESPONSIVE ================= */

        @media (max-width: 700px) {

          .page {
            padding: 20px 10px;
          }

          .letterheadText {
            padding: 25px 22px 22px;
          }

          .importSection,
          .searchSection,
          .results {
            margin-left: 22px;
            margin-right: 22px;
          }

          .searchGrid {
            grid-template-columns: 1fr;
          }

          .searchHeader {
            gap: 15px;
          }

          .searchInfo {
            flex-direction: column;
            align-items: flex-start;
            gap: 6px;
          }

          .resultsHeader {
            align-items: flex-start;
            gap: 12px;
          }

          .tableWrap {
            max-height: 500px;
          }

        }

      `}</style>
    </div>
  );
}