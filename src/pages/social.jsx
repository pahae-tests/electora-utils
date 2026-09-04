import { useState } from "react";
import * as XLSX from "xlsx";

export default function SocialPage() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
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

  const findColumn = (columns, names) => {
    for (const name of names) {
      const found = columns.find(
        (column) =>
          normalize(column) === normalize(name)
      );

      if (found) {
        return found;
      }
    }

    return null;
  };

  // =========================================================
  // CRÉER UN EXCEL
  // =========================================================

  const downloadExcel = (data, filename) => {
    const worksheet = XLSX.utils.json_to_sheet(data);

    // Largeur des colonnes
    worksheet["!cols"] = [
      { wch: 25 }, // nom
      { wch: 25 }, // prénom
      { wch: 16 }, // cin
      { wch: 20 }, // téléphone
      { wch: 28 }, // bureau de vote
      { wch: 25 }, // parrain
      { wch: 20 }, // service social
    ];

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Électeurs"
    );

    XLSX.writeFile(
      workbook,
      filename
    );
  };

  // =========================================================
  // TRAITEMENT DU FICHIER
  // =========================================================

  const processFile = async (selectedFile) => {
    if (!selectedFile) return;

    setLoading(true);
    setError("");
    setResult(null);

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

      if (!workbook.SheetNames.length) {
        throw new Error(
          "Le fichier ne contient aucune feuille."
        );
      }

      // Première feuille
      const sheetName =
        workbook.SheetNames[0];

      const worksheet =
        workbook.Sheets[sheetName];

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
      // Colonnes présentes
      // -----------------------------------------------------

      const columns =
        Object.keys(rows[0]);

      // -----------------------------------------------------
      // Détection des colonnes
      // -----------------------------------------------------

      const nomColumn = findColumn(
        columns,
        [
          "Nom",
          "nom",
        ]
      );

      const prenomColumn = findColumn(
        columns,
        [
          "Prénom",
          "Prenom",
          "prénom",
          "prenom",
        ]
      );

      const cinColumn = findColumn(
        columns,
        [
          "numeroCIN",
          "Numéro CIN",
          "Numero CIN",
          "numero CIN",
          "CIN",
          "cin",
        ]
      );

      const telephoneColumn = findColumn(
        columns,
        [
          "Téléphone",
          "Telephone",
          "telephone",
          "Tel",
          "tel",
          "Téléphone mobile",
        ]
      );

      const bureauColumn = findColumn(
        columns,
        [
          "Bureau de vote",
          "Bureau vote",
          "bureauVote",
          "bureau vote",
          "bureau de vote",
        ]
      );

      const parrainColumn = findColumn(
        columns,
        [
          "Parrain",
          "parrain",
          "parrainNom",
          "Parrain Nom",
          "Nom parrain",
        ]
      );

      const ajouteParColumn = findColumn(
        columns,
        [
          "Ajouté par",
          "Ajoute par",
          "ajoutePar",
          "Ajout par",
          "ajouté par",
          "ajoute par",
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
          "numeroCIN / CIN"
        );
      }

      if (!telephoneColumn) {
        missing.push("Téléphone");
      }

      if (!bureauColumn) {
        missing.push(
          "Bureau de vote"
        );
      }

      if (!parrainColumn) {
        missing.push("Parrain");
      }

      if (!ajouteParColumn) {
        missing.push("Ajouté par");
      }

      if (missing.length > 0) {
        throw new Error(
          `Colonnes obligatoires introuvables :\n\n${missing.join(
            "\n"
          )}\n\nColonnes présentes dans votre fichier :\n${columns.join(
            ", "
          )}`
        );
      }

      // =====================================================
      // SÉPARATION ET SUPPRESSION DES DOUBLONS
      // =====================================================

      const khalidRows = [];
      const otherRows = [];

      // CIN déjà rencontrés
      const khalidCins = new Set();
      const otherCins = new Set();

      let duplicatesKhalid = 0;
      let duplicatesOther = 0;

      rows.forEach((row) => {

        // ---------------------------------------------------
        // Qui a ajouté l'électeur ?
        // ---------------------------------------------------

        const ajoutePar =
          normalize(
            row[ajouteParColumn]
          );

        const isKhalid =
          ajoutePar ===
          normalize("Khalid Touzani");

        // ---------------------------------------------------
        // CIN
        // ---------------------------------------------------

        const cin =
          String(
            row[cinColumn] ?? ""
          ).trim();

        const normalizedCin =
          normalize(cin);

        // ---------------------------------------------------
        // Construction de la ligne finale
        // ---------------------------------------------------

        const formattedRow = {
          nom:
            row[nomColumn] ?? "",

          prénom:
            row[prenomColumn] ?? "",

          cin:
            cin,

          téléphone:
            row[telephoneColumn] ?? "",

          "bureau de vote":
            row[bureauColumn] ?? "",

          parrain:
            row[parrainColumn] ?? "",

          "service social":
            isKhalid
              ? "Moltaqa"
              : "Social",
        };

        // ===================================================
        // KHALID TOUZANI
        // ===================================================

        if (isKhalid) {

          // Si le CIN est renseigné
          // et déjà présent => doublon
          if (
            normalizedCin &&
            khalidCins.has(
              normalizedCin
            )
          ) {
            duplicatesKhalid++;
            return;
          }

          if (normalizedCin) {
            khalidCins.add(
              normalizedCin
            );
          }

          khalidRows.push(
            formattedRow
          );

          return;
        }

        // ===================================================
        // AUTRES
        // ===================================================

        if (
          normalizedCin &&
          otherCins.has(
            normalizedCin
          )
        ) {
          duplicatesOther++;
          return;
        }

        if (normalizedCin) {
          otherCins.add(
            normalizedCin
          );
        }

        otherRows.push(
          formattedRow
        );
      });

      // =====================================================
      // RÉSULTAT
      // =====================================================

      setFile(selectedFile);

      setResult({
        total: rows.length,

        khalid: khalidRows,

        other: otherRows,

        duplicatesKhalid,

        duplicatesOther,

        totalDuplicates:
          duplicatesKhalid +
          duplicatesOther,

        sheetName,

        columns: {
          nom: nomColumn,
          prenom: prenomColumn,
          cin: cinColumn,
          telephone:
            telephoneColumn,
          bureau:
            bureauColumn,
          parrain:
            parrainColumn,
          ajoutePar:
            ajouteParColumn,
        },
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
  // TÉLÉCHARGER KHALID
  // =========================================================

  const downloadKhalid = () => {
    if (!result) return;

    downloadExcel(
      result.khalid,
      "electeurs_Khalid_Touzani.xlsx"
    );
  };

  // =========================================================
  // TÉLÉCHARGER AUTRES
  // =========================================================

  const downloadOther = () => {
    if (!result) return;

    downloadExcel(
      result.other,
      "electeurs_autres.xlsx"
    );
  };

  // =========================================================
  // TÉLÉCHARGER LES DEUX
  // =========================================================

  const downloadBoth = () => {
    if (!result) return;

    downloadExcel(
      result.khalid,
      "electeurs_Khalid_Touzani.xlsx"
    );

    setTimeout(() => {
      downloadExcel(
        result.other,
        "electeurs_autres.xlsx"
      );
    }, 500);
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
              Traitement électoral
            </span>

            <h1>
              Répartition — Service social
            </h1>

            <p>
              Importez un fichier Excel pour séparer
              les électeurs ajoutés par Khalid Touzani
              des autres électeurs et supprimer les doublons.
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
            RÉSULTATS
        ================================================= */}

        {result && (

          <section className="resultsSection">

            <div className="resultHeader">

              <div>

                <span className="sectionLabel">
                  Résultat
                </span>

                <h2>
                  Fichiers prêts
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
                  lignes dans le fichier
                </small>

              </div>

              <div className="statCard khalid">

                <span>
                  Khalid Touzani
                </span>

                <strong>
                  {result.khalid.length.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

                <small>
                  Moltaqa
                </small>

              </div>

              <div className="statCard other">

                <span>
                  Autres
                </span>

                <strong>
                  {result.other.length.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

                <small>
                  Social
                </small>

              </div>

            </div>

            {/* =================================================
                DOUBLONS
            ================================================= */}

            <div className="duplicatesBox">

              <div className="duplicatesHeader">

                <div>

                  <span className="sectionLabel">
                    Nettoyage
                  </span>

                  <h3>
                    Doublons supprimés
                  </h3>

                </div>

                <strong>
                  {result.totalDuplicates.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="duplicatesRows">

                <div>

                  <span>
                    Doublons — Khalid Touzani
                  </span>

                  <strong>
                    {result.duplicatesKhalid.toLocaleString(
                      "fr-FR"
                    )}
                  </strong>

                </div>

                <div>

                  <span>
                    Doublons — Autres
                  </span>

                  <strong>
                    {result.duplicatesOther.toLocaleString(
                      "fr-FR"
                    )}
                  </strong>

                </div>

              </div>

              <p className="duplicateNote">
                Les doublons sont identifiés à partir
                du CIN. Une seule ligne est conservée
                pour chaque CIN dans chaque fichier.
              </p>

            </div>

            {/* =================================================
                VÉRIFICATION
            ================================================= */}

            <div className="verification">

              <div className="verificationTitle">
                Vérification
              </div>

              <div className="verificationRow">

                <span>
                  Khalid Touzani après nettoyage
                </span>

                <strong>
                  {result.khalid.length.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  Autres après nettoyage
                </span>

                <strong>
                  {result.other.length.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationTotal">

                <span>
                  Total après nettoyage
                </span>

                <strong>
                  {(
                    result.khalid.length +
                    result.other.length
                  ).toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

            </div>

            {/* =================================================
                FICHIERS
            ================================================= */}

            <div className="filesGrid">

              {/* =================================================
                  KHALID
              ================================================= */}

              <div className="outputCard">

                <div className="outputTop">

                  <div className="excelIcon">
                    XLS
                  </div>

                  <div>

                    <h3>
                      Khalid Touzani
                    </h3>

                    <p>
                      Service social :{" "}
                      <strong>
                        Moltaqa
                      </strong>
                    </p>

                  </div>

                </div>

                <div className="outputCount">

                  {result.khalid.length.toLocaleString(
                    "fr-FR"
                  )}

                  {" "}

                  électeur
                  {result.khalid.length !== 1
                    ? "s"
                    : ""}

                </div>

                <div className="outputInfo">
                  {result.duplicatesKhalid > 0
                    ? `${result.duplicatesKhalid.toLocaleString(
                        "fr-FR"
                      )} doublon${
                        result.duplicatesKhalid !== 1
                          ? "s"
                          : ""
                      } supprimé${
                        result.duplicatesKhalid !== 1
                          ? "s"
                          : ""
                      }`
                    : "Aucun doublon"}
                </div>

                <button
                  type="button"
                  className="downloadBtn"
                  onClick={downloadKhalid}
                >
                  ↓ Télécharger
                </button>

              </div>

              {/* =================================================
                  AUTRES
              ================================================= */}

              <div className="outputCard">

                <div className="outputTop">

                  <div className="excelIcon">
                    XLS
                  </div>

                  <div>

                    <h3>
                      Autres électeurs
                    </h3>

                    <p>
                      Service social :{" "}
                      <strong>
                        Social
                      </strong>
                    </p>

                  </div>

                </div>

                <div className="outputCount">

                  {result.other.length.toLocaleString(
                    "fr-FR"
                  )}

                  {" "}

                  électeur
                  {result.other.length !== 1
                    ? "s"
                    : ""}

                </div>

                <div className="outputInfo">
                  {result.duplicatesOther > 0
                    ? `${result.duplicatesOther.toLocaleString(
                        "fr-FR"
                      )} doublon${
                        result.duplicatesOther !== 1
                          ? "s"
                          : ""
                      } supprimé${
                        result.duplicatesOther !== 1
                          ? "s"
                          : ""
                      }`
                    : "Aucun doublon"}
                </div>

                <button
                  type="button"
                  className="downloadBtn"
                  onClick={downloadOther}
                >
                  ↓ Télécharger
                </button>

              </div>

            </div>

            {/* =================================================
                DOWNLOAD BOTH
            ================================================= */}

            <button
              type="button"
              className="downloadAllBtn"
              onClick={downloadBoth}
            >
              ↓ Télécharger les deux fichiers
            </button>

            {/* =================================================
                COLONNES
            ================================================= */}

            <div className="columnsInfo">

              <div className="columnsTitle">
                Colonnes des fichiers générés
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
                  téléphone
                </span>

                <span>
                  bureau de vote
                </span>

                <span>
                  parrain
                </span>

                <span>
                  service social
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
                générer les deux fichiers sociaux.
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

        /* ================= HEADER ================= */

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

        /* ================= SECTIONS ================= */

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

        /* ================= DROP ZONE ================= */

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

        /* ================= ERROR ================= */

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

        /* ================= STATS ================= */

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

        .statCard.other {
          border-top-color: #3f7657;
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

        /* ================= DUPLICATES ================= */

        .duplicatesBox {
          margin-top: 20px;
          border: 1px solid #e0ddd4;
          background: #fbfaf8;
          padding: 18px 20px;
        }

        .duplicatesHeader {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding-bottom: 14px;
        }

        .duplicatesHeader h3 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 17px;
          margin: 4px 0 0;
          font-weight: 600;
        }

        .duplicatesHeader > strong {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 28px;
          color: #1f3a5f;
        }

        .duplicatesRows {
          border-top: 1px solid #e0ddd4;
        }

        .duplicatesRows > div {
          display: flex;
          justify-content: space-between;
          padding: 10px 0;
          border-bottom: 1px solid #ede9e0;
          font-size: 12px;
          color: #6b6459;
        }

        .duplicatesRows strong {
          color: #1e2124;
        }

        .duplicateNote {
          margin: 12px 0 0;
          font-size: 11px;
          color: #8a8378;
          line-height: 1.5;
        }

        /* ================= VERIFICATION ================= */

        .verification {
          margin-top: 20px;
          border-top: 2px solid #1f3a5f;
          border-bottom: 1px solid #e0ddd4;
        }

        .verificationTitle {
          padding: 12px 0 9px;
          font-size: 11px;
          text-transform: uppercase;
          letter-spacing: 0.06em;
          color: #96723a;
          font-weight: 600;
        }

        .verificationRow,
        .verificationTotal {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 9px 0;
          border-top: 1px solid #ede9e0;
          font-size: 12px;
          color: #6b6459;
        }

        .verificationRow strong {
          color: #1e2124;
        }

        .verificationTotal {
          border-top-color: #c9c3b7;
          color: #1e2124;
          font-weight: 600;
        }

        .verificationTotal strong {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 18px;
          color: #1f3a5f;
        }

        /* ================= FILES ================= */

        .filesGrid {
          display: grid;
          grid-template-columns:
            repeat(2, 1fr);
          gap: 14px;
          margin-top: 22px;
        }

        .outputCard {
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
          flex-shrink: 0;
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
        }

        .outputCard p strong {
          color: #1f3a5f;
        }

        .outputCount {
          margin: 20px 0 6px;
          color: #4a4740;
          font-size: 13px;
        }

        .outputInfo {
          font-size: 11px;
          color: #8a8378;
          min-height: 17px;
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

        /* ================= DOWNLOAD ALL ================= */

        .downloadAllBtn {
          width: 100%;
          margin-top: 14px;
          padding: 12px;
          border: none;
          background: #1f3a5f;
          color: #ffffff;
          border-radius: 4px;
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
        }

        .downloadAllBtn:hover {
          background: #172d49;
        }

        /* ================= COLUMNS ================= */

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

        /* ================= RESET ================= */

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

        /* ================= EMPTY ================= */

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

        /* ================= RESPONSIVE ================= */

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

          .statsGrid,
          .filesGrid {
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

        }

      `}</style>

    </div>
  );
}