import { useState } from "react";
import * as XLSX from "xlsx";

export default function VerifyPage() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);

  // =========================================================
  // NORMALISATION
  // =========================================================

  const normalize = (value) => {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  };

  // =========================================================
  // RECHERCHE D'UNE COLONNE
  // =========================================================

  const findColumn = (columns, possibleNames) => {
    for (const name of possibleNames) {
      const normalizedName = normalize(name);

      const found = columns.find(
        (column) =>
          normalize(column) === normalizedName
      );

      if (found) {
        return found;
      }
    }

    return null;
  };

  // =========================================================
  // VÉRIFIER SI LE BUREAU DE VOTE EST VIDE
  // =========================================================

  const isEmptyBureau = (value) => {
    if (
      value === null ||
      value === undefined
    ) {
      return true;
    }

    const text = String(value)
      .trim()
      .toLowerCase();

    if (text === "") {
      return true;
    }

    const emptyValues = [
      "aucun",
      "aucune",
      "غير معروف",
      "غيرمعروف",
      "inconnu",
      "inconnue",
      "unknown",
      "null",
      "undefined",
      "n/a",
      "na",
      "-"
    ];

    return emptyValues.includes(text);
  };

  // =========================================================
  // TRAITEMENT DU FICHIER
  // =========================================================

  const processFile = async (selectedFile) => {
    if (!selectedFile) return;

    setLoading(true);
    setError("");
    setResult(null);
    setFile(selectedFile);

    try {
      // -----------------------------------------------------
      // Vérification extension
      // -----------------------------------------------------

      const fileName =
        selectedFile.name.toLowerCase();

      if (
        !fileName.endsWith(".xlsx") &&
        !fileName.endsWith(".xls") &&
        !fileName.endsWith(".csv")
      ) {
        throw new Error(
          "Veuillez sélectionner un fichier Excel (.xlsx, .xls ou .csv)."
        );
      }

      // -----------------------------------------------------
      // Lecture
      // -----------------------------------------------------

      const buffer =
        await selectedFile.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
      });

      if (
        !workbook.SheetNames ||
        workbook.SheetNames.length === 0
      ) {
        throw new Error(
          "Le fichier ne contient aucune feuille."
        );
      }

      const sheetName =
        workbook.SheetNames[0];

      const worksheet =
        workbook.Sheets[sheetName];

      // -----------------------------------------------------
      // Conversion en objets
      // -----------------------------------------------------

      const rows =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            defval: "",
            raw: false,
          }
        );

      if (!rows.length) {
        throw new Error(
          "La feuille Excel est vide."
        );
      }

      // -----------------------------------------------------
      // Colonnes
      // -----------------------------------------------------

      const columns =
        Object.keys(rows[0]);

      // -----------------------------------------------------
      // Détection des colonnes
      // -----------------------------------------------------

      const nomColumn =
        findColumn(
          columns,
          [
            "Nom",
            "nom",
          ]
        );

      const prenomColumn =
        findColumn(
          columns,
          [
            "Prénom",
            "Prenom",
            "prénom",
            "prenom",
          ]
        );

      const cinColumn =
        findColumn(
          columns,
          [
            "CIN",
            "cin",
            "numeroCIN",
            "Numéro CIN",
            "Numero CIN",
            "numero CIN",
          ]
        );

      const parrainColumn =
        findColumn(
          columns,
          [
            "Parrain",
            "parrain",
            "parrainNom",
            "Parrain Nom",
            "Nom parrain",
          ]
        );

      const ajouteParColumn =
        findColumn(
          columns,
          [
            "Ajouté par",
            "Ajoute par",
            "ajouté par",
            "ajoute par",
            "Ajout par",
            "AjoutéPar",
            "ajoutePar",
          ]
        );

      const bureauColumn =
        findColumn(
          columns,
          [
            "Bureau de vote",
            "Bureau vote",
            "bureau de vote",
            "bureau vote",
            "bureauVote",
            "BureauVote",
            "bureau_vote",
          ]
        );

      // -----------------------------------------------------
      // Vérification
      // -----------------------------------------------------

      const missing = [];

      if (!nomColumn) {
        missing.push("Nom");
      }

      if (!prenomColumn) {
        missing.push("Prénom");
      }

      if (!cinColumn) {
        missing.push(
          "CIN / numeroCIN"
        );
      }

      if (!parrainColumn) {
        missing.push("Parrain");
      }

      if (!ajouteParColumn) {
        missing.push("Ajouté par");
      }

      if (!bureauColumn) {
        missing.push("Bureau de vote");
      }

      if (missing.length > 0) {
        throw new Error(
          `Colonnes obligatoires introuvables :\n\n${missing.join(
            "\n"
          )}\n\nColonnes détectées dans votre fichier :\n${columns.join(
            ", "
          )}`
        );
      }

      // =====================================================
      // FILTRAGE
      // =====================================================

      const filteredRows =
        rows.filter((row) => {

          const ajoutePar =
            normalize(
              row[ajouteParColumn]
            );

          const isKhalid =
            ajoutePar ===
            normalize(
              "Khalid Touzani"
            );

          const bureauVide =
            isEmptyBureau(
              row[bureauColumn]
            );

          return (
            isKhalid &&
            bureauVide
          );
        });

      // =====================================================
      // CRÉATION DU NOUVEAU FICHIER
      // =====================================================
      //
      // IMPORTANT :
      // Le fichier final contient UNIQUEMENT :
      //
      // nom | prénom | cin | parrain
      //
      // =====================================================

      const outputRows =
        filteredRows.map((row) => ({
          nom:
            row[nomColumn] ?? "",

          prénom:
            row[prenomColumn] ?? "",

          cin:
            row[cinColumn] ?? "",

          parrain:
            row[parrainColumn] ?? "",
        }));

      // -----------------------------------------------------
      // Création de la feuille
      // -----------------------------------------------------

      const outputWorksheet =
        XLSX.utils.json_to_sheet(
          outputRows,
          {
            header: [
              "nom",
              "prénom",
              "cin",
              "parrain",
            ],
          }
        );

      // -----------------------------------------------------
      // Largeur colonnes
      // -----------------------------------------------------

      outputWorksheet["!cols"] = [
        { wch: 25 },
        { wch: 25 },
        { wch: 18 },
        { wch: 30 },
      ];

      // -----------------------------------------------------
      // Création workbook
      // -----------------------------------------------------

      const outputWorkbook =
        XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        outputWorksheet,
        "Vérification"
      );

      // =====================================================
      // TÉLÉCHARGEMENT AUTOMATIQUE
      // =====================================================

      const outputFileName =
        "Khalid_Touzani_bureau_vote_non_reconnu.xlsx";

      XLSX.writeFile(
        outputWorkbook,
        outputFileName
      );

      // =====================================================
      // RÉSULTAT
      // =====================================================

      setResult({
        total:
          rows.length,

        filtered:
          filteredRows.length,

        ajouteParColumn,

        bureauColumn,

        sheetName,

        fileName:
          outputFileName,
      });

    } catch (err) {
      console.error(err);

      setError(
        err.message ||
          "Une erreur est survenue lors du traitement du fichier."
      );

      setFile(null);

    } finally {
      setLoading(false);
    }
  };

  // =========================================================
  // CHANGEMENT FICHIER
  // =========================================================

  const handleFileChange = (event) => {
    const selectedFile =
      event.target.files?.[0];

    if (selectedFile) {
      processFile(selectedFile);
    }
  };

  // =========================================================
  // DRAG & DROP
  // =========================================================

  const handleDrop = (event) => {
    event.preventDefault();

    setDragging(false);

    const droppedFile =
      event.dataTransfer.files?.[0];

    if (droppedFile) {
      processFile(droppedFile);
    }
  };

  // =========================================================
  // RESET
  // =========================================================

  const reset = () => {
    setFile(null);
    setResult(null);
    setError("");
    setLoading(false);
  };

  // =========================================================
  // RENDU
  // =========================================================

  return (
    <div className="page">

      <div className="sheet">

        {/* =================================================
            HEADER
        ================================================= */}

        <header className="letterhead">

          <div className="letterheadBar" />

          <div className="letterheadText">

            <span className="eyebrow">
              Vérification électorale
            </span>

            <h1>
              Bureaux de vote non reconnus
            </h1>

            <p>
              Importez un fichier Excel pour
              extraire les électeurs ajoutés par
              Khalid Touzani dont le bureau de vote
              est vide ou non reconnu.
            </p>

          </div>

        </header>

        {/* =================================================
            IMPORTATION
        ================================================= */}

        <section className="uploadSection">

          <div className="sectionHeading">

            <div>

              <span className="sectionLabel">
                Importation
              </span>

              <h2>
                Sélectionner le fichier Excel
              </h2>

            </div>

          </div>

          <label
            className={`dropZone ${
              dragging
                ? "dragging"
                : ""
            }`}
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() =>
              setDragging(false)
            }
            onDrop={handleDrop}
          >

            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
            />

            <div className="uploadIcon">
              ↑
            </div>

            <div className="uploadTitle">

              {loading
                ? "Traitement du fichier..."
                : file
                ? file.name
                : "Déposez votre fichier Excel ici"}

            </div>

            <div className="uploadText">
              ou cliquez pour sélectionner un fichier
            </div>

            <div className="uploadFormats">
              XLSX · XLS · CSV
            </div>

          </label>

          {/* =================================================
              ERROR
          ================================================= */}

          {error && (

            <div className="error">

              <div className="errorTitle">
                Impossible de traiter le fichier
              </div>

              <div className="errorText">
                {error}
              </div>

            </div>

          )}

        </section>

        {/* =================================================
            RÉSULTAT
        ================================================= */}

        {result && (

          <section className="resultsSection">

            <div className="resultHeader">

              <div>

                <span className="sectionLabel">
                  Résultat
                </span>

                <h2>
                  Extraction terminée
                </h2>

              </div>

              <button
                type="button"
                className="resetBtn"
                onClick={reset}
              >
                Nouveau fichier
              </button>

            </div>

            {/* =================================================
                STATISTIQUES
            ================================================= */}

            <div className="statsGrid">

              <div className="statCard">

                <span>
                  Total original
                </span>

                <strong>
                  {result.total.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

                <small>
                  lignes analysées
                </small>

              </div>

              <div className="statCard khalid">

                <span>
                  Khalid Touzani
                </span>

                <strong>
                  {result.filtered.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

                <small>
                  bureau de vote vide
                </small>

              </div>

              <div className="statCard">

                <span>
                  Non sélectionnées
                </span>

                <strong>
                  {(
                    result.total -
                    result.filtered
                  ).toLocaleString(
                    "fr-FR"
                  )}
                </strong>

                <small>
                  lignes restantes
                </small>

              </div>

            </div>

            {/* =================================================
                CONDITIONS
            ================================================= */}

            <div className="conditionsBox">

              <div className="conditionTitle">
                Conditions appliquées
              </div>

              <div className="conditionRows">

                <div className="conditionRow">

                  <span>
                    Ajouté par
                  </span>

                  <strong>
                    Khalid Touzani
                  </strong>

                </div>

                <div className="conditionRow">

                  <span>
                    Bureau de vote
                  </span>

                  <strong>
                    Vide / Aucun / Non reconnu
                  </strong>

                </div>

              </div>

            </div>

            {/* =================================================
                FICHIER
            ================================================= */}

            <div className="outputCard">

              <div className="outputTop">

                <div className="excelIcon">
                  XLS
                </div>

                <div>

                  <h3>
                    Fichier généré
                  </h3>

                  <p>
                    {result.fileName}
                  </p>

                </div>

              </div>

              <div className="outputCount">

                {result.filtered.toLocaleString(
                  "fr-FR"
                )}

                {" "}

                ligne
                {result.filtered !== 1
                  ? "s"
                  : ""}

              </div>

              <div className="outputInfo">
                Colonnes : nom, prénom, cin, parrain
              </div>

              <button
                type="button"
                className="downloadBtn"
                onClick={() => {

                  if (file) {
                    processFile(file);
                  }

                }}
              >
                ↓ Générer et télécharger à nouveau
              </button>

            </div>

            {/* =================================================
                COLONNES
            ================================================= */}

            <div className="columnsInfo">

              <div className="columnsTitle">
                Colonnes du fichier généré
              </div>

              <div className="columnsList">

                <span>
                  nom
                </span>

                <span>
                  prénom
                </span>

                <span>
                  cin
                </span>

                <span>
                  parrain
                </span>

              </div>

            </div>

          </section>

        )}

        {/* =================================================
            EMPTY STATE
        ================================================= */}

        {!result &&
          !loading &&
          !error && (

            <div className="emptyState">

              <div className="emptyIcon">
                XLS
              </div>

              <h3>
                Aucun fichier analysé
              </h3>

              <p>
                Importez votre fichier Excel pour
                lancer automatiquement l'extraction.
              </p>

            </div>

          )}

      </div>

      {/* =====================================================
          STYLE
      ===================================================== */}

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
          max-width: 1050px;
          background: #ffffff;
          border: 1px solid #e0ddd4;
          border-radius: 4px;
          overflow: hidden;
        }

        .letterhead {
          display: flex;
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
          max-width: 700px;
          line-height: 1.6;
        }

        .uploadSection,
        .resultsSection {
          margin: 30px 36px;
        }

        .sectionHeading,
        .resultHeader {
          display: flex;
          justify-content: space-between;
          align-items: flex-end;
          margin-bottom: 16px;
        }

        .sectionLabel {
          color: #96723a;
          font-size: 11px;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: 0.06em;
        }

        h2 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 19px;
          font-weight: 600;
          margin: 4px 0 0;
          color: #16191c;
        }

        .dropZone {
          min-height: 190px;
          border: 1px dashed #c9c3b7;
          background: #fbfaf8;
          display: flex;
          flex-direction: column;
          justify-content: center;
          align-items: center;
          text-align: center;
          cursor: pointer;
          transition: all 0.15s ease;
          border-radius: 4px;
        }

        .dropZone:hover,
        .dropZone.dragging {
          border-color: #1f3a5f;
          background: #f5f7f9;
        }

        .dropZone input {
          display: none;
        }

        .uploadIcon {
          width: 45px;
          height: 45px;
          border: 1px solid #d7d2c6;
          background: #ffffff;
          display: flex;
          justify-content: center;
          align-items: center;
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 22px;
          color: #1f3a5f;
          margin-bottom: 12px;
        }

        .uploadTitle {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 17px;
          color: #1e2124;
          margin-bottom: 5px;
        }

        .uploadText {
          font-size: 12px;
          color: #8a8378;
        }

        .uploadFormats {
          margin-top: 12px;
          font-size: 10px;
          letter-spacing: 0.08em;
          color: #96723a;
          font-weight: 600;
        }

        .error {
          margin-top: 16px;
          padding: 14px 16px;
          border: 1px solid #e3b8b8;
          background: #fbeeee;
          border-radius: 4px;
          color: #8a2d2d;
        }

        .errorTitle {
          font-size: 13px;
          font-weight: 600;
          margin-bottom: 4px;
        }

        .errorText {
          font-size: 12px;
          white-space: pre-line;
          line-height: 1.5;
        }

        .statsGrid {
          display: grid;
          grid-template-columns:
            repeat(3, 1fr);
          gap: 14px;
        }

        .statCard {
          border: 1px solid #e0ddd4;
          border-top: 3px solid #1f3a5f;
          padding: 18px 20px;
        }

        .statCard.khalid {
          border-top-color: #96723a;
        }

        .statCard span {
          display: block;
          font-size: 11px;
          color: #6b6459;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          font-weight: 600;
        }

        .statCard strong {
          display: block;
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 32px;
          color: #16191c;
          margin-top: 10px;
        }

        .statCard small {
          color: #8a8378;
          font-size: 11px;
        }

        .conditionsBox {
          margin-top: 20px;
          border: 1px solid #e0ddd4;
          background: #fbfaf8;
          padding: 18px 20px;
        }

        .conditionTitle {
          font-size: 11px;
          color: #96723a;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          margin-bottom: 10px;
        }

        .conditionRow {
          display: flex;
          justify-content: space-between;
          padding: 10px 0;
          border-top: 1px solid #ede9e0;
          font-size: 12px;
        }

        .conditionRow span {
          color: #6b6459;
        }

        .conditionRow strong {
          color: #1f3a5f;
        }

        .outputCard {
          margin-top: 20px;
          border: 1px solid #e0ddd4;
          padding: 20px;
          background: #fbfaf8;
        }

        .outputTop {
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .excelIcon {
          width: 42px;
          height: 42px;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #ffffff;
          border: 1px solid #d7d2c6;
          color: #3f7657;
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 11px;
          font-weight: 600;
        }

        .outputCard h3 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 16px;
          margin: 0 0 4px;
          font-weight: 600;
        }

        .outputCard p {
          margin: 0;
          font-size: 11px;
          color: #8a8378;
          word-break: break-all;
        }

        .outputCount {
          margin: 20px 0 6px;
          color: #4a4740;
          font-size: 13px;
        }

        .outputInfo {
          font-size: 11px;
          color: #8a8378;
          margin-bottom: 14px;
        }

        .downloadBtn {
          width: 100%;
          padding: 10px;
          border: 1px solid #1f3a5f;
          background: transparent;
          color: #1f3a5f;
          border-radius: 4px;
          font-size: 12px;
          font-weight: 600;
          cursor: pointer;
        }

        .downloadBtn:hover {
          background: #eef2f6;
        }

        .columnsInfo {
          margin-top: 22px;
          padding-top: 18px;
          border-top: 1px solid #e0ddd4;
        }

        .columnsTitle {
          font-size: 11px;
          color: #6b6459;
          font-weight: 600;
          margin-bottom: 10px;
        }

        .columnsList {
          display: flex;
          flex-wrap: wrap;
          gap: 7px;
        }

        .columnsList span {
          padding: 5px 9px;
          border: 1px solid #ddd7cb;
          background: #f5f2ec;
          color: #6b6459;
          font-size: 11px;
          border-radius: 3px;
        }

        .resetBtn {
          background: transparent;
          border: 1px solid #1f3a5f;
          color: #1f3a5f;
          border-radius: 4px;
          padding: 8px 14px;
          font-size: 12px;
          font-weight: 600;
          cursor: pointer;
        }

        .resetBtn:hover {
          background: #eef2f6;
        }

        .emptyState {
          margin: 0 36px 36px;
          padding: 55px 20px;
          border-top: 1px solid #e0ddd4;
          text-align: center;
        }

        .emptyIcon {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          height: 42px;
          padding: 0 12px;
          border: 1px solid #d7d2c6;
          color: #1f3a5f;
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 14px;
          margin-bottom: 12px;
        }

        .emptyState h3 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 17px;
          font-weight: 600;
          margin: 0 0 6px;
        }

        .emptyState p {
          margin: 0;
          color: #8a8378;
          font-size: 13px;
        }

        @media (max-width: 700px) {

          .page {
            padding: 20px 10px;
          }

          .letterheadText {
            padding: 25px 22px;
          }

          .uploadSection,
          .resultsSection {
            margin-left: 22px;
            margin-right: 22px;
          }

          .statsGrid {
            grid-template-columns: 1fr;
          }

          .resultHeader {
            align-items: flex-start;
            flex-direction: column;
            gap: 12px;
          }

          .resetBtn {
            width: 100%;
          }

          .conditionRow {
            flex-direction: column;
            gap: 4px;
          }

        }

      `}</style>

    </div>
  );
}