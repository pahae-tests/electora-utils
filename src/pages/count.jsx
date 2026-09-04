import { useState } from "react";
import * as XLSX from "xlsx";

export default function CountPage() {
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
  // RECONNAÎTRE LE NOM DE LA COLONNE BUREAU DE VOTE
  // =========================================================

  const isBureauVoteColumn = (columnName) => {
    const name = normalize(columnName)
      .replace(/[_-]/g, " ");

    const candidates = [
      "bureau vote",
      "bureau de vote",
      "bureauvote",
      "bureau",
      "مكتب التصويت",
      "مكتب الاقتراع",
      "bureau_vote",
    ];

    return candidates.some((candidate) => {
      const normalizedCandidate = normalize(candidate)
        .replace(/[_-]/g, " ");

      return (
        name === normalizedCandidate ||
        name.includes(normalizedCandidate)
      );
    });
  };

  // =========================================================
  // RECONNAÎTRE LE NOM DE LA COLONNE AGENT
  // =========================================================

  const isAgentColumn = (columnName) => {
    const name = normalize(columnName)
      .replace(/[_-]/g, " ");

    const candidates = [
      "agent",
      "nom agent",
      "agent saisie",
      "agent de saisie",
      "utilisateur",
      "user",
      "created by",
      "createdby",
      "created_by",
      "ajoute par",
      "ajouté par",
      "ajoutepar",
      "ajoute_par",
      "operateur",
      "operateur de saisie",
      "opérateur",
      "opérateur de saisie",
      "saisi par",
      "saisie par",
      "inserted by",
      "insertedby",
      "responsable",
      "nom utilisateur",
    ];

    return candidates.some((candidate) => {
      const normalizedCandidate = normalize(candidate)
        .replace(/[_-]/g, " ");

      return (
        name === normalizedCandidate ||
        name.includes(normalizedCandidate)
      );
    });
  };

  // =========================================================
  // VALEUR = BUREAU NON RECONNU ?
  // =========================================================

  const isUnknownBureau = (value) => {
    const normalized = normalize(value);

    // Valeur vide
    if (!normalized) {
      return true;
    }

    // Valeurs françaises
    const frenchValues = [
      "aucun",
      "aucune",
      "inconnu",
      "inconnue",
      "non reconnu",
      "non reconnue",
      "non renseigne",
      "non renseigné",
      "non defini",
      "non défini",
      "indisponible",
      "neant",
      "néant",
      "n/a",
      "na",
      "null",
      "undefined",
      "-",
      "--",
    ];

    if (frenchValues.includes(normalized)) {
      return true;
    }

    // Valeurs anglaises
    const englishValues = [
      "unknown",
      "none",
      "not known",
      "not recognized",
      "unrecognized",
      "undefined",
      "null",
      "n/a",
      "na",
    ];

    if (englishValues.includes(normalized)) {
      return true;
    }

    // Valeurs arabes
    const arabicValues = [
      "غير معروف",
      "غير معروفة",
      "غير محدد",
      "غير محددة",
      "غير متوفر",
      "غير متوفرة",
      "لا يوجد",
      "لا توجد",
      "مجهول",
      "مجهولة",
      "غير موجود",
      "غير موجودة",
      "لا شيء",
    ];

    if (arabicValues.includes(normalized)) {
      return true;
    }

    return false;
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
      const fileName = selectedFile.name.toLowerCase();

      if (
        !fileName.endsWith(".xlsx") &&
        !fileName.endsWith(".xls") &&
        !fileName.endsWith(".csv")
      ) {
        throw new Error(
          "Veuillez sélectionner un fichier Excel (.xlsx, .xls ou .csv)."
        );
      }

      const buffer = await selectedFile.arrayBuffer();

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
      const firstSheetName =
        workbook.SheetNames[0];

      const worksheet =
        workbook.Sheets[firstSheetName];

      // Convertir en objets
      const data = XLSX.utils.sheet_to_json(
        worksheet,
        {
          defval: "",
          raw: false,
        }
      );

      if (!data.length) {
        throw new Error(
          "La feuille Excel est vide."
        );
      }

      // =====================================================
      // DÉTECTION DES COLONNES
      // =====================================================

      const columns = Object.keys(data[0]);

      // -----------------------------------------------------
      // Colonne BUREAU DE VOTE
      // -----------------------------------------------------

      let bureauColumn = columns.find(
        (column) =>
          isBureauVoteColumn(column)
      );

      // Recherche plus souple
      if (!bureauColumn) {
        bureauColumn = columns.find(
          (column) => {
            const normalized = normalize(column);

            return (
              normalized.includes("bureau") &&
              normalized.includes("vote")
            );
          }
        );
      }

      if (!bureauColumn) {
        throw new Error(
          `Impossible de trouver la colonne "Bureau de vote".

Colonnes trouvées :
${columns.join(", ")}`
        );
      }

      // -----------------------------------------------------
      // Colonne AGENT
      // -----------------------------------------------------

      let agentColumn = columns.find(
        (column) =>
          isAgentColumn(column)
      );

      // Recherche encore plus souple
      if (!agentColumn) {
        agentColumn = columns.find(
          (column) => {
            const normalized = normalize(column);

            return (
              normalized.includes("agent") ||
              normalized.includes("utilisateur") ||
              normalized.includes("user") ||
              normalized.includes("ajoute") ||
              normalized.includes("ajout") ||
              normalized.includes("saisi") ||
              normalized.includes("saisie") ||
              normalized.includes("operateur") ||
              normalized.includes("responsable")
            );
          }
        );
      }

      if (!agentColumn) {
        throw new Error(
          `Impossible de trouver la colonne contenant le nom de l'agent.

Pour calculer Moltaqa et Social, le fichier doit contenir une colonne indiquant l'agent ayant ajouté l'électeur.

Colonnes trouvées :
${columns.join(", ")}`
        );
      }

      // =====================================================
      // COMPTAGE
      // =====================================================

      const total = data.length;

      // -----------------------------
      // Bureaux
      // -----------------------------

      let nonReconnu = 0;
      let reconnu = 0;

      // -----------------------------
      // Moltaqa
      // -----------------------------

      let moltaqa = 0;
      let moltaqaReconnu = 0;
      let moltaqaNonReconnu = 0;

      // -----------------------------
      // Social
      // -----------------------------

      let social = 0;
      let socialReconnu = 0;
      let socialNonReconnu = 0;

      const khalidTouzani =
        normalize("Khalid Touzani");

      // =====================================================
      // PARCOURIR LES ÉLECTEURS
      // =====================================================

      data.forEach((row) => {

        const bureau =
          row[bureauColumn];

        const agent =
          normalize(row[agentColumn]);

        const isMoltaqa =
          agent === khalidTouzani;

        const isRecognized =
          !isUnknownBureau(bureau);

        // ===================================================
        // BUREAU GLOBAL
        // ===================================================

        if (isRecognized) {
          reconnu++;
        } else {
          nonReconnu++;
        }

        // ===================================================
        // MOLTAQA / SOCIAL
        // ===================================================

        if (isMoltaqa) {

          // Total Moltaqa
          moltaqa++;

          // Moltaqa reconnu / non reconnu
          if (isRecognized) {
            moltaqaReconnu++;
          } else {
            moltaqaNonReconnu++;
          }

        } else {

          // Total Social
          social++;

          // Social reconnu / non reconnu
          if (isRecognized) {
            socialReconnu++;
          } else {
            socialNonReconnu++;
          }
        }
      });

      // =====================================================
      // QUELQUES VALEURS POUR APERÇU
      // =====================================================

      const unknownExamples = [];

      data.forEach((row) => {
        const value = row[bureauColumn];

        if (
          isUnknownBureau(value) &&
          value !== ""
        ) {
          const stringValue =
            String(value).trim();

          if (
            stringValue &&
            !unknownExamples.includes(
              stringValue
            )
          ) {
            unknownExamples.push(
              stringValue
            );
          }
        }
      });

      // =====================================================
      // ENREGISTRER LE RÉSULTAT
      // =====================================================

      setResult({

        // Global
        total,
        nonReconnu,
        reconnu,

        // Moltaqa
        moltaqa,
        moltaqaReconnu,
        moltaqaNonReconnu,

        // Social
        social,
        socialReconnu,
        socialNonReconnu,

        // Colonnes
        bureauColumn,
        agentColumn,

        sheetName: firstSheetName,
        columns,

        unknownExamples:
          unknownExamples.slice(0, 10),
      });

      setFile(selectedFile);

    } catch (err) {
      console.error(err);

      setError(
        err.message ||
          "Une erreur est survenue lors de la lecture du fichier."
      );

      setFile(null);

    } finally {
      setLoading(false);
    }
  };

  // =========================================================
  // INPUT
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
  // POURCENTAGE
  // =========================================================

  const percentage = (value, total = null) => {

    const denominator =
      total !== null
        ? total
        : result?.total;

    if (!denominator || denominator === 0) {
      return "0.00";
    }

    return (
      (value / denominator) *
      100
    ).toFixed(2);
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
              Analyse électorale
            </span>

            <h1>
              Comptage des électeurs
            </h1>

            <p>
              Importez un fichier Excel pour obtenir
              automatiquement les statistiques relatives
              aux électeurs, aux bureaux de vote et aux
              agents de saisie.
            </p>

          </div>

        </header>

        {/* =================================================
            UPLOAD
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
                ? "Analyse du fichier..."
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
                Impossible d'analyser le fichier
              </div>

              <div className="errorText">
                {error}
              </div>

            </div>
          )}

        </section>

        {/* =================================================
            RESULTATS
        ================================================= */}

        {result && (

          <section className="resultsSection">

            <div className="resultHeader">

              <div>

                <span className="sectionLabel">
                  Résultat de l'analyse
                </span>

                <h2>
                  Statistiques des électeurs
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
                FILE INFO
            ================================================= */}

            <div className="fileInfo">

              <div className="fileInfoItem">

                <span>
                  Fichier
                </span>

                <strong>
                  {file?.name}
                </strong>

              </div>

              <div className="fileInfoItem">

                <span>
                  Feuille
                </span>

                <strong>
                  {result.sheetName}
                </strong>

              </div>

              <div className="fileInfoItem">

                <span>
                  Colonne bureau
                </span>

                <strong>
                  {result.bureauColumn}
                </strong>

              </div>

              <div className="fileInfoItem">

                <span>
                  Colonne agent
                </span>

                <strong>
                  {result.agentColumn}
                </strong>

              </div>

            </div>

            {/* =================================================
                CARTES PRINCIPALES
            ================================================= */}

            <div className="cards">

              {/* TOTAL */}

              <div className="card totalCard">

                <div className="cardTop">

                  <span className="cardLabel">
                    Total des électeurs
                  </span>

                  <span className="cardIcon">
                    #
                  </span>

                </div>

                <div className="cardNumber">
                  {result.total.toLocaleString(
                    "fr-FR"
                  )}
                </div>

                <div className="cardDescription">
                  Nombre total de lignes
                  d'électeurs dans le fichier.
                </div>

              </div>

              {/* MOLTAQA */}

              <div className="card moltaqaCard">

                <div className="cardTop">

                  <span className="cardLabel">
                    Moltaqa
                  </span>

                  <span className="cardIcon">
                    M
                  </span>

                </div>

                <div className="cardNumber">
                  {result.moltaqa.toLocaleString(
                    "fr-FR"
                  )}
                </div>

                <div className="cardDescription">
                  Électeurs ajoutés par
                  <strong> Khalid Touzani</strong>.
                </div>

                <div className="percentage">
                  {percentage(result.moltaqa)}%
                </div>

              </div>

              {/* SOCIAL */}

              <div className="card socialCard">

                <div className="cardTop">

                  <span className="cardLabel">
                    Social
                  </span>

                  <span className="cardIcon">
                    S
                  </span>

                </div>

                <div className="cardNumber">
                  {result.social.toLocaleString(
                    "fr-FR"
                  )}
                </div>

                <div className="cardDescription">
                  Électeurs qui n'ont pas été
                  ajoutés par Khalid Touzani.
                </div>

                <div className="percentage">
                  {percentage(result.social)}%
                </div>

              </div>

              {/* BUREAU NON RECONNU */}

              <div className="card unknownCard">

                <div className="cardTop">

                  <span className="cardLabel">
                    Bureau non reconnu
                  </span>

                  <span className="cardIcon">
                    ?
                  </span>

                </div>

                <div className="cardNumber">
                  {result.nonReconnu.toLocaleString(
                    "fr-FR"
                  )}
                </div>

                <div className="cardDescription">
                  Aucun bureau, bureau inconnu
                  ou valeur similaire.
                </div>

                <div className="percentage">
                  {percentage(result.nonReconnu)}%
                </div>

              </div>

              {/* BUREAU RECONNU */}

              <div className="card recognizedCard">

                <div className="cardTop">

                  <span className="cardLabel">
                    Bureau reconnu
                  </span>

                  <span className="cardIcon">
                    ✓
                  </span>

                </div>

                <div className="cardNumber">
                  {result.reconnu.toLocaleString(
                    "fr-FR"
                  )}
                </div>

                <div className="cardDescription">
                  Électeurs possédant une
                  valeur de bureau de vote.
                </div>

                <div className="percentage">
                  {percentage(result.reconnu)}%
                </div>

              </div>

            </div>

            {/* =================================================
                DÉTAIL MOLTAQA
            ================================================= */}

            <div className="detailSection moltaqaDetails">

              <div className="detailHeader">

                <div>
                  <span className="sectionLabel">
                    Moltaqa
                  </span>

                  <h3>
                    Détail des électeurs de Khalid Touzani
                  </h3>
                </div>

                <div className="detailTotal">
                  {result.moltaqa.toLocaleString(
                    "fr-FR"
                  )}
                  <span> électeurs</span>
                </div>

              </div>

              <div className="detailCards">

                <div className="detailCard recognizedDetail">

                  <div className="detailCardLabel">
                    Bureau reconnu
                  </div>

                  <div className="detailCardNumber">
                    {result.moltaqaReconnu.toLocaleString(
                      "fr-FR"
                    )}
                  </div>

                  <div className="detailCardPercentage">
                    {percentage(
                      result.moltaqaReconnu,
                      result.moltaqa
                    )}%
                    <span>
                      {" "}des Moltaqa
                    </span>
                  </div>

                </div>

                <div className="detailCard unknownDetail">

                  <div className="detailCardLabel">
                    Bureau non reconnu
                  </div>

                  <div className="detailCardNumber">
                    {result.moltaqaNonReconnu.toLocaleString(
                      "fr-FR"
                    )}
                  </div>

                  <div className="detailCardPercentage">
                    {percentage(
                      result.moltaqaNonReconnu,
                      result.moltaqa
                    )}%
                    <span>
                      {" "}des Moltaqa
                    </span>
                  </div>

                </div>

              </div>

            </div>

            {/* =================================================
                DÉTAIL SOCIAL
            ================================================= */}

            <div className="detailSection socialDetails">

              <div className="detailHeader">

                <div>
                  <span className="sectionLabel">
                    Social
                  </span>

                  <h3>
                    Détail des électeurs des autres agents
                  </h3>
                </div>

                <div className="detailTotal">
                  {result.social.toLocaleString(
                    "fr-FR"
                  )}
                  <span> électeurs</span>
                </div>

              </div>

              <div className="detailCards">

                <div className="detailCard recognizedDetail">

                  <div className="detailCardLabel">
                    Bureau reconnu
                  </div>

                  <div className="detailCardNumber">
                    {result.socialReconnu.toLocaleString(
                      "fr-FR"
                    )}
                  </div>

                  <div className="detailCardPercentage">
                    {percentage(
                      result.socialReconnu,
                      result.social
                    )}%
                    <span>
                      {" "}des Social
                    </span>
                  </div>

                </div>

                <div className="detailCard unknownDetail">

                  <div className="detailCardLabel">
                    Bureau non reconnu
                  </div>

                  <div className="detailCardNumber">
                    {result.socialNonReconnu.toLocaleString(
                      "fr-FR"
                    )}
                  </div>

                  <div className="detailCardPercentage">
                    {percentage(
                      result.socialNonReconnu,
                      result.social
                    )}%
                    <span>
                      {" "}des Social
                    </span>
                  </div>

                </div>

              </div>

            </div>

            {/* =================================================
                VÉRIFICATION MOLTAQA / SOCIAL
            ================================================= */}

            <div className="verification">

              <div className="verificationTitle">
                Vérification Moltaqa / Social
              </div>

              <div className="verificationRow">

                <span>
                  Total
                </span>

                <strong>
                  {result.total.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  Moltaqa — Khalid Touzani
                </span>

                <strong>
                  {result.moltaqa.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  └─ Bureau reconnu
                </span>

                <strong>
                  {result.moltaqaReconnu.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  └─ Bureau non reconnu
                </span>

                <strong>
                  {result.moltaqaNonReconnu.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  Social — autres agents
                </span>

                <strong>
                  {result.social.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  └─ Bureau reconnu
                </span>

                <strong>
                  {result.socialReconnu.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  └─ Bureau non reconnu
                </span>

                <strong>
                  {result.socialNonReconnu.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationTotal">

                <span>
                  Moltaqa + Social
                </span>

                <strong>
                  {(
                    result.moltaqa +
                    result.social
                  ).toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

            </div>

            {/* =================================================
                VÉRIFICATION BUREAUX
            ================================================= */}

            <div className="verification">

              <div className="verificationTitle">
                Vérification des bureaux
              </div>

              <div className="verificationRow">

                <span>
                  Total
                </span>

                <strong>
                  {result.total.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  Bureau non reconnu
                </span>

                <strong>
                  {result.nonReconnu.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationRow">

                <span>
                  Bureau reconnu
                </span>

                <strong>
                  {result.reconnu.toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

              <div className="verificationTotal">

                <span>
                  Non reconnu + reconnu
                </span>

                <strong>
                  {(
                    result.nonReconnu +
                    result.reconnu
                  ).toLocaleString(
                    "fr-FR"
                  )}
                </strong>

              </div>

            </div>

            {/* =================================================
                VALEURS NON RECONNUES
            ================================================= */}

            {result.unknownExamples.length >
              0 && (

              <div className="examples">

                <div className="examplesTitle">
                  Valeurs considérées comme
                  « non reconnues »
                </div>

                <div className="chips">

                  {result.unknownExamples.map(
                    (value, index) => (
                      <span
                        className="chip"
                        key={`${value}-${index}`}
                      >
                        {value}
                      </span>
                    )
                  )}

                </div>

              </div>
            )}

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
              commencer le comptage.
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
          max-width: 1250px;
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
          max-width: 750px;
          line-height: 1.5;
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
          position: relative;
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

        /* ================= FILE INFO ================= */

        .fileInfo {
          display: grid;
          grid-template-columns:
            1.4fr
            1fr
            1fr
            1fr;
          border: 1px solid #e0ddd4;
          background: #fbfaf8;
          margin-bottom: 18px;
        }

        .fileInfoItem {
          padding: 12px 15px;
          border-right: 1px solid #e0ddd4;
          display: flex;
          flex-direction: column;
          gap: 4px;
          min-width: 0;
        }

        .fileInfoItem:last-child {
          border-right: none;
        }

        .fileInfoItem span {
          font-size: 10px;
          color: #8a8378;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }

        .fileInfoItem strong {
          font-size: 12px;
          color: #4a4740;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }

        /* ================= CARDS ================= */

        .cards {
          display: grid;
          grid-template-columns:
            repeat(5, 1fr);
          gap: 14px;
        }

        .card {
          border: 1px solid #e0ddd4;
          padding: 20px;
          min-height: 155px;
          position: relative;
          background: #ffffff;
        }

        .totalCard {
          border-top: 3px solid #1f3a5f;
        }

        .moltaqaCard {
          border-top: 3px solid #7c3aed;
        }

        .socialCard {
          border-top: 3px solid #2563eb;
        }

        .unknownCard {
          border-top: 3px solid #96723a;
        }

        .recognizedCard {
          border-top: 3px solid #3f7657;
        }

        .cardTop {
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .cardLabel {
          font-size: 11px;
          color: #6b6459;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          font-weight: 600;
        }

        .cardIcon {
          width: 26px;
          height: 26px;
          border: 1px solid #d7d2c6;
          display: flex;
          align-items: center;
          justify-content: center;
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          color: #96723a;
        }

        .moltaqaCard .cardIcon {
          color: #7c3aed;
        }

        .socialCard .cardIcon {
          color: #2563eb;
        }

        .recognizedCard .cardIcon {
          color: #3f7657;
        }

        .cardNumber {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 34px;
          font-weight: 600;
          color: #16191c;
          margin-top: 17px;
          line-height: 1;
        }

        .cardDescription {
          margin-top: 12px;
          color: #8a8378;
          font-size: 11px;
          line-height: 1.45;
          padding-right: 35px;
        }

        .cardDescription strong {
          color: #4a4740;
        }

        .percentage {
          position: absolute;
          right: 20px;
          bottom: 19px;
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 15px;
          color: #1f3a5f;
        }

        .moltaqaCard .percentage {
          color: #7c3aed;
        }

        .socialCard .percentage {
          color: #2563eb;
        }

        .unknownCard .percentage {
          color: #96723a;
        }

        .recognizedCard .percentage {
          color: #3f7657;
        }

        /* ================= DETAIL SECTIONS ================= */

        .detailSection {
          margin-top: 22px;
          border: 1px solid #e0ddd4;
          padding: 20px;
          background: #fbfaf8;
        }

        .moltaqaDetails {
          border-top: 3px solid #7c3aed;
        }

        .socialDetails {
          border-top: 3px solid #2563eb;
        }

        .detailHeader {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 16px;
        }

        .detailHeader h3 {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 16px;
          font-weight: 600;
          margin: 5px 0 0;
          color: #16191c;
        }

        .detailTotal {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 24px;
          font-weight: 600;
          color: #16191c;
        }

        .detailTotal span {
          font-family:
            -apple-system,
            BlinkMacSystemFont,
            "Segoe UI",
            Roboto,
            sans-serif;
          font-size: 11px;
          color: #8a8378;
          font-weight: 400;
        }

        .detailCards {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 14px;
        }

        .detailCard {
          background: #ffffff;
          border: 1px solid #e0ddd4;
          padding: 18px;
          position: relative;
        }

        .recognizedDetail {
          border-left: 3px solid #3f7657;
        }

        .unknownDetail {
          border-left: 3px solid #96723a;
        }

        .detailCardLabel {
          font-size: 11px;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          color: #6b6459;
          font-weight: 600;
        }

        .detailCardNumber {
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 30px;
          font-weight: 600;
          margin-top: 12px;
          color: #16191c;
        }

        .detailCardPercentage {
          margin-top: 7px;
          font-family:
            Georgia,
            "Times New Roman",
            serif;
          font-size: 14px;
          color: #3f7657;
        }

        .unknownDetail .detailCardPercentage {
          color: #96723a;
        }

        .detailCardPercentage span {
          font-family:
            -apple-system,
            BlinkMacSystemFont,
            "Segoe UI",
            Roboto,
            sans-serif;
          font-size: 11px;
          color: #8a8378;
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
          font-weight: 600;
        }

        .verificationTotal {
          margin-top: 3px;
          border-top: 1px solid #c9c3b7;
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

        /* ================= EXAMPLES ================= */

        .examples {
          margin-top: 22px;
          padding-top: 18px;
          border-top: 1px solid #e0ddd4;
        }

        .examplesTitle {
          font-size: 12px;
          font-weight: 600;
          color: #4a4740;
          margin-bottom: 10px;
        }

        .chips {
          display: flex;
          flex-wrap: wrap;
          gap: 7px;
        }

        .chip {
          padding: 5px 9px;
          background: #f5f2ec;
          border: 1px solid #ddd7cb;
          color: #6b6459;
          font-size: 11px;
          border-radius: 3px;
        }

        /* ================= BUTTON ================= */

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

        @media (max-width: 1100px) {

          .cards {
            grid-template-columns:
              repeat(3, 1fr);
          }

          .fileInfo {
            grid-template-columns:
              repeat(2, 1fr);
          }

          .fileInfoItem:nth-child(2) {
            border-right: none;
          }

          .fileInfoItem:nth-child(-n + 2) {
            border-bottom: 1px solid #e0ddd4;
          }
        }

        @media (max-width: 750px) {

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

          .cards {
            grid-template-columns: 1fr;
          }

          .detailCards {
            grid-template-columns: 1fr;
          }

          .detailHeader {
            align-items: flex-start;
            flex-direction: column;
            gap: 10px;
          }

          .fileInfo {
            grid-template-columns: 1fr;
          }

          .fileInfoItem {
            border-right: none !important;
            border-bottom: 1px solid #e0ddd4;
          }

          .fileInfoItem:last-child {
            border-bottom: none;
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