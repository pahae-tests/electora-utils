import { useState } from "react";
import * as XLSX from "xlsx";

export default function MergePage() {
    const [files, setFiles] = useState([]);
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
    // VÉRIFICATION DES FICHIERS
    // =========================================================

    const isExcelFile = (file) => {
        const name = file.name.toLowerCase();

        return (
            name.endsWith(".xlsx") ||
            name.endsWith(".xls") ||
            name.endsWith(".csv")
        );
    };

    // =========================================================
    // AJOUT DES FICHIERS
    // =========================================================

    const handleFiles = (selectedFiles) => {
        setError("");

        const selected = Array.from(selectedFiles);

        if (selected.length === 0) {
            return;
        }

        // Vérifier les extensions
        const invalidFiles = selected.filter(
            (file) => !isExcelFile(file)
        );

        if (invalidFiles.length > 0) {
            setError(
                "Seuls les fichiers Excel (.xlsx, .xls ou .csv) sont acceptés."
            );
            return;
        }

        // Nombre maximum de fichiers
        if (files.length + selected.length > 3) {
            setError(
                "Vous pouvez sélectionner au maximum 3 fichiers Excel au total."
            );
            return;
        }

        // Ajouter les nouveaux fichiers aux fichiers existants
        setFiles((previousFiles) => [
            ...previousFiles,
            ...selected,
        ]);

        setResult(null);
    };

    // =========================================================
    // INPUT FILE
    // =========================================================

    const handleFileChange = (event) => {
        handleFiles(event.target.files);
    };

    // =========================================================
    // DRAG & DROP
    // =========================================================

    const handleDrop = (event) => {
        event.preventDefault();
        setDragging(false);

        handleFiles(event.dataTransfer.files);
    };

    // =========================================================
    // SUPPRIMER UN FICHIER
    // =========================================================

    const removeFile = (index) => {
        const newFiles = files.filter(
            (_, i) => i !== index
        );

        setFiles(newFiles);
        setResult(null);
        setError("");
    };

    // =========================================================
    // DÉPLACER UN FICHIER VERS LE HAUT
    // =========================================================

    const moveUp = (index) => {
        if (index === 0) return;

        const newFiles = [...files];

        [
            newFiles[index - 1],
            newFiles[index],
        ] = [
                newFiles[index],
                newFiles[index - 1],
            ];

        setFiles(newFiles);
        setResult(null);
    };

    // =========================================================
    // DÉPLACER UN FICHIER VERS LE BAS
    // =========================================================

    const moveDown = (index) => {
        if (index === files.length - 1) return;

        const newFiles = [...files];

        [
            newFiles[index],
            newFiles[index + 1],
        ] = [
                newFiles[index + 1],
                newFiles[index],
            ];

        setFiles(newFiles);
        setResult(null);
    };

    // =========================================================
    // LIRE UN FICHIER
    // =========================================================

    const readExcel = async (file) => {
        const buffer =
            await file.arrayBuffer();

        const workbook =
            XLSX.read(buffer, {
                type: "array",
                cellDates: true,
            });

        if (
            !workbook.SheetNames ||
            workbook.SheetNames.length === 0
        ) {
            throw new Error(
                `Le fichier "${file.name}" ne contient aucune feuille.`
            );
        }

        // Première feuille
        const sheetName =
            workbook.SheetNames[0];

        const worksheet =
            workbook.Sheets[sheetName];

        // Lecture sous forme de tableau
        const rows =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    header: 1,
                    defval: "",
                    raw: false,
                }
            );

        if (!rows.length) {
            return {
                fileName: file.name,
                sheetName,
                headers: [],
                data: [],
            };
        }

        // Première ligne = en-têtes
        const headers =
            rows[0].map((header) =>
                String(header ?? "").trim()
            );

        // Supprimer :
        // 1. les lignes complètement vides
        // 2. les lignes de Total
        // 3. les lignes où seule "service social" est remplie

        const serviceSocialIndex = headers.findIndex(
            (header) =>
                normalize(header) === "service social"
        );

        const data = rows
            .slice(1)
            .filter((row) => {
                // -------------------------------------------------------
                // 1. Ligne complètement vide
                // -------------------------------------------------------
                const isEmpty = !row.some(
                    (cell) =>
                        String(cell ?? "").trim() !== ""
                );

                if (isEmpty) return false;

                // -------------------------------------------------------
                // 2. Ligne de Total
                // -------------------------------------------------------
                const firstCells = row
                    .slice(0, 3)
                    .map((cell) =>
                        String(cell ?? "")
                            .trim()
                            .toLowerCase()
                            .normalize("NFD")
                            .replace(/[\u0300-\u036f]/g, "")
                    )
                    .join(" ");

                const isTotal =
                    /\btotal\b/.test(firstCells) ||
                    /\btotal general\b/.test(firstCells) ||
                    /\bsomme\b/.test(firstCells);

                if (isTotal) return false;

                // -------------------------------------------------------
                // 3. Seule la colonne "service social" est remplie
                // -------------------------------------------------------
                if (serviceSocialIndex !== -1) {
                    const serviceSocialValue = String(
                        row[serviceSocialIndex] ?? ""
                    ).trim();

                    const otherColumnsFilled = row.some(
                        (cell, index) =>
                            index !== serviceSocialIndex &&
                            String(cell ?? "").trim() !== ""
                    );

                    const onlyServiceSocialFilled =
                        serviceSocialValue !== "" &&
                        !otherColumnsFilled;

                    if (onlyServiceSocialFilled) {
                        return false;
                    }
                }

                return true;
            });

        return {
            fileName: file.name,
            sheetName,
            headers,
            data,
        };
    };

    // =========================================================
    // FUSION
    // =========================================================

    const mergeFiles = async () => {
        if (files.length < 2) {
            setError(
                "Veuillez sélectionner au moins 2 fichiers."
            );
            return;
        }

        setLoading(true);
        setError("");
        setResult(null);

        try {
            // -----------------------------------------------------
            // Lecture de tous les fichiers
            // -----------------------------------------------------

            const parsedFiles = [];

            for (const file of files) {
                const parsed =
                    await readExcel(file);

                parsedFiles.push(parsed);
            }

            // -----------------------------------------------------
            // Vérifier qu'il existe au moins une colonne
            // -----------------------------------------------------

            const validFiles =
                parsedFiles.filter(
                    (file) =>
                        file.headers.length > 0
                );

            if (!validFiles.length) {
                throw new Error(
                    "Les fichiers sélectionnés sont vides."
                );
            }

            // -----------------------------------------------------
            // Colonnes finales
            // -----------------------------------------------------
            //
            // On prend les colonnes du premier fichier comme
            // référence.
            //
            // Si les autres fichiers possèdent des colonnes
            // supplémentaires, elles sont également ajoutées.
            //
            // -----------------------------------------------------

            const finalHeaders = [];

            parsedFiles.forEach((file) => {
                file.headers.forEach((header) => {
                    const exists =
                        finalHeaders.some(
                            (existing) =>
                                normalize(existing) ===
                                normalize(header)
                        );

                    if (
                        header &&
                        !exists
                    ) {
                        finalHeaders.push(
                            header
                        );
                    }
                });
            });

            if (!finalHeaders.length) {
                throw new Error(
                    "Aucune colonne n'a été trouvée dans les fichiers."
                );
            }

            // =====================================================
            // CONSTRUCTION DES LIGNES
            // =====================================================

            const mergedRows = [];

            parsedFiles.forEach(
                (file, fileIndex) => {

                    file.data.forEach(
                        (row) => {

                            const newRow = {};

                            finalHeaders.forEach(
                                (header) => {

                                    // Trouver la colonne correspondante
                                    // dans le fichier actuel
                                    const columnIndex =
                                        file.headers.findIndex(
                                            (fileHeader) =>
                                                normalize(
                                                    fileHeader
                                                ) ===
                                                normalize(
                                                    header
                                                )
                                        );

                                    if (
                                        columnIndex !==
                                        -1
                                    ) {
                                        newRow[header] =
                                            row[
                                            columnIndex
                                            ] ?? "";
                                    } else {
                                        newRow[header] =
                                            "";
                                    }
                                }
                            );

                            mergedRows.push(
                                newRow
                            );
                        }
                    );
                }
            );

            // =====================================================
            // CRÉATION DU FICHIER EXCEL
            // =====================================================

            const worksheet =
                XLSX.utils.json_to_sheet(
                    mergedRows,
                    {
                        header:
                            finalHeaders,
                    }
                );

            // =====================================================
            // LARGEUR DES COLONNES
            // =====================================================

            worksheet["!cols"] =
                finalHeaders.map(
                    (header) => {

                        let maxLength =
                            header.length;

                        for (
                            let i = 0;
                            i <
                            Math.min(
                                mergedRows.length,
                                1000
                            );
                            i++
                        ) {
                            const value =
                                mergedRows[i][
                                header
                                ];

                            const length =
                                String(
                                    value ?? ""
                                ).length;

                            if (
                                length >
                                maxLength
                            ) {
                                maxLength =
                                    length;
                            }
                        }

                        return {
                            wch: Math.min(
                                Math.max(
                                    maxLength + 2,
                                    10
                                ),
                                40
                            ),
                        };
                    }
                );

            // =====================================================
            // WORKBOOK
            // =====================================================

            const workbook =
                XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "Fusion"
            );

            // =====================================================
            // TÉLÉCHARGEMENT
            // =====================================================
            const firstFileName =
                parsedFiles[0]?.fileName || "fichiers_fusionnes.xlsx";

            const outputFileName =
                firstFileName.replace(/\.(xlsx|xls|csv)$/i, "") + ".xlsx";

            XLSX.writeFile(
                workbook,
                outputFileName
            );

            // =====================================================
            // STATISTIQUES
            // =====================================================

            const fileStats =
                parsedFiles.map(
                    (file) => ({
                        name:
                            file.fileName,
                        rows:
                            file.data.length,
                    })
                );

            setResult({
                totalFiles:
                    parsedFiles.length,

                totalRows:
                    mergedRows.length,

                headers:
                    finalHeaders,

                files:
                    fileStats,

                outputFileName,
            });

        } catch (err) {
            console.error(err);

            setError(
                err.message ||
                "Une erreur est survenue pendant la fusion."
            );

        } finally {
            setLoading(false);
        }
    };

    // =========================================================
    // RESET
    // =========================================================

    const reset = () => {
        setFiles([]);
        setResult(null);
        setError("");
        setLoading(false);
    };

    // =========================================================
    // FORMAT TAILLE
    // =========================================================

    const formatNumber = (number) => {
        return Number(
            number || 0
        ).toLocaleString("fr-FR");
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
                            Gestion des fichiers
                        </span>

                        <h1>
                            Fusionner des fichiers Excel
                        </h1>

                        <p>
                            Sélectionnez deux ou trois fichiers
                            Excel pour les fusionner dans un seul
                            fichier, en conservant l'ordre de
                            sélection.
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
                                01 — Fichiers
                            </span>

                            <h2>
                                Sélectionner 2 ou 3 fichiers
                            </h2>

                        </div>

                        {files.length > 0 && (
                            <span className="fileCounter">
                                {files.length} / 3
                            </span>
                        )}

                    </div>

                    <label
                        className={`dropZone ${dragging
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
                            multiple
                            onChange={handleFileChange}
                        />

                        <div className="uploadIcon">
                            +
                        </div>

                        <div className="uploadTitle">
                            Déposez un fichier Excel ici
                        </div>

                        <div className="uploadText">
                            ou cliquez pour ajouter un fichier
                        </div>

                        <div className="uploadFormats">
                            XLSX · XLS · CSV
                        </div>

                    </label>

                    {/* =================================================
              ERREUR
          ================================================= */}

                    {error && (

                        <div className="error">

                            <div className="errorTitle">
                                Impossible de continuer
                            </div>

                            <div className="errorText">
                                {error}
                            </div>

                        </div>

                    )}

                </section>

                {/* =================================================
            LISTE DES FICHIERS
        ================================================= */}

                {files.length > 0 && (

                    <section className="filesSection">

                        <div className="sectionHeading">

                            <div>

                                <span className="sectionLabel">
                                    02 — Ordre de fusion
                                </span>

                                <h2>
                                    Fichiers sélectionnés
                                </h2>

                            </div>

                        </div>

                        <div className="fileList">

                            {files.map(
                                (file, index) => (

                                    <div
                                        className="fileItem"
                                        key={`${file.name}-${index}`}
                                    >

                                        <div className="fileNumber">
                                            {index + 1}
                                        </div>

                                        <div className="excelIcon">
                                            XLS
                                        </div>

                                        <div className="fileDetails">

                                            <strong>
                                                {file.name}
                                            </strong>

                                            <span>
                                                {(
                                                    file.size /
                                                    1024
                                                ).toFixed(1)}{" "}
                                                Ko
                                            </span>

                                        </div>

                                        <div className="fileActions">

                                            <button
                                                type="button"
                                                onClick={() =>
                                                    moveUp(index)
                                                }
                                                disabled={
                                                    index === 0
                                                }
                                                title="Déplacer vers le haut"
                                            >
                                                ↑
                                            </button>

                                            <button
                                                type="button"
                                                onClick={() =>
                                                    moveDown(index)
                                                }
                                                disabled={
                                                    index ===
                                                    files.length - 1
                                                }
                                                title="Déplacer vers le bas"
                                            >
                                                ↓
                                            </button>

                                            <button
                                                type="button"
                                                className="remove"
                                                onClick={() =>
                                                    removeFile(index)
                                                }
                                                title="Supprimer"
                                            >
                                                ×
                                            </button>

                                        </div>

                                    </div>

                                )
                            )}

                        </div>

                        <div className="orderInfo">

                            <span>
                                Ordre de fusion :
                            </span>

                            <strong>
                                {files
                                    .map(
                                        (_, index) =>
                                            `Excel ${index + 1}`
                                    )
                                    .join(" → ")}
                            </strong>

                        </div>

                        {/* =================================================
                BOUTON FUSION
            ================================================= */}

                        <button
                            type="button"
                            className="mergeBtn"
                            onClick={mergeFiles}
                            disabled={
                                files.length < 2 ||
                                loading
                            }
                        >

                            {loading
                                ? "Fusion en cours..."
                                : "↓ Fusionner et télécharger"}

                        </button>

                    </section>

                )}

                {/* =================================================
            RÉSULTAT
        ================================================= */}

                {result && (

                    <section className="resultsSection">

                        <div className="resultHeader">

                            <div>

                                <span className="sectionLabel">
                                    03 — Résultat
                                </span>

                                <h2>
                                    Fusion terminée
                                </h2>

                            </div>

                            <button
                                type="button"
                                className="resetBtn"
                                onClick={reset}
                            >
                                Nouvelle fusion
                            </button>

                        </div>

                        {/* =================================================
                STATS
            ================================================= */}

                        <div className="statsGrid">

                            <div className="statCard">

                                <span>
                                    Fichiers fusionnés
                                </span>

                                <strong>
                                    {result.totalFiles}
                                </strong>

                                <small>
                                    fichiers Excel
                                </small>

                            </div>

                            <div className="statCard">

                                <span>
                                    Total des lignes
                                </span>

                                <strong>
                                    {formatNumber(
                                        result.totalRows
                                    )}
                                </strong>

                                <small>
                                    après fusion
                                </small>

                            </div>

                            <div className="statCard">

                                <span>
                                    Colonnes
                                </span>

                                <strong>
                                    {result.headers.length}
                                </strong>

                                <small>
                                    colonnes finales
                                </small>

                            </div>

                        </div>

                        {/* =================================================
                DÉTAIL DES FICHIERS
            ================================================= */}

                        <div className="detailsBox">

                            <div className="detailsTitle">
                                Répartition des lignes
                            </div>

                            {result.files.map(
                                (file, index) => (

                                    <div
                                        className="detailRow"
                                        key={`${file.name}-${index}`}
                                    >

                                        <div className="detailLeft">

                                            <span className="detailNumber">
                                                {index + 1}
                                            </span>

                                            <span>
                                                {file.name}
                                            </span>

                                        </div>

                                        <strong>
                                            {formatNumber(
                                                file.rows
                                            )}{" "}
                                            ligne
                                            {file.rows !== 1
                                                ? "s"
                                                : ""}
                                        </strong>

                                    </div>

                                )
                            )}

                        </div>

                        {/* =================================================
                FICHIER GÉNÉRÉ
            ================================================= */}

                        <div className="outputCard">

                            <div className="outputTop">

                                <div className="excelIcon">
                                    XLS
                                </div>

                                <div>

                                    <h3>
                                        Fichier fusionné
                                    </h3>

                                    <p>
                                        {result.outputFileName}
                                    </p>

                                </div>

                            </div>

                            <div className="outputCount">

                                {formatNumber(
                                    result.totalRows
                                )}{" "}
                                lignes

                            </div>

                            <div className="outputInfo">
                                Le fichier a été téléchargé automatiquement.
                            </div>

                        </div>

                        {/* =================================================
                COLONNES
            ================================================= */}

                        <div className="columnsInfo">

                            <div className="columnsTitle">
                                Colonnes du fichier final
                            </div>

                            <div className="columnsList">

                                {result.headers.map(
                                    (header) => (

                                        <span
                                            key={header}
                                        >
                                            {header}
                                        </span>

                                    )
                                )}

                            </div>

                        </div>

                    </section>

                )}

                {/* =================================================
            EMPTY STATE
        ================================================= */}

                {files.length === 0 &&
                    !result &&
                    !loading &&
                    !error && (

                        <div className="emptyState">

                            <div className="emptyIcon">
                                XLS
                            </div>

                            <h3>
                                Aucun fichier sélectionné
                            </h3>

                            <p>
                                Sélectionnez 2 ou 3 fichiers Excel
                                pour commencer.
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
        .filesSection,
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

        .fileCounter {
          color: #1f3a5f;
          font-size: 12px;
          font-weight: 600;
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
          font-size: 24px;
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

        /* ================= FILE LIST ================= */

        .fileList {
          border: 1px solid #e0ddd4;
        }

        .fileItem {
          display: flex;
          align-items: center;
          padding: 13px 14px;
          border-bottom: 1px solid #ede9e0;
          gap: 12px;
        }

        .fileItem:last-child {
          border-bottom: none;
        }

        .fileNumber {
          width: 25px;
          height: 25px;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #1f3a5f;
          color: #ffffff;
          font-size: 11px;
          font-weight: 600;
          border-radius: 50%;
          flex-shrink: 0;
        }

        .excelIcon {
          width: 40px;
          height: 40px;
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
          font-size: 10px;
          font-weight: 600;
          flex-shrink: 0;
        }

        .fileDetails {
          flex: 1;
          min-width: 0;
        }

        .fileDetails strong {
          display: block;
          font-size: 13px;
          font-weight: 600;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }

        .fileDetails span {
          display: block;
          margin-top: 3px;
          font-size: 10px;
          color: #8a8378;
        }

        .fileActions {
          display: flex;
          gap: 4px;
        }

        .fileActions button {
          width: 29px;
          height: 29px;
          border: 1px solid #d7d2c6;
          background: #ffffff;
          color: #1f3a5f;
          cursor: pointer;
          border-radius: 3px;
          font-size: 14px;
        }

        .fileActions button:hover:not(:disabled) {
          background: #eef2f6;
        }

        .fileActions button:disabled {
          color: #c9c3b7;
          cursor: not-allowed;
        }

        .fileActions button.remove {
          color: #a33d3d;
        }

        .fileActions button.remove:hover {
          background: #fbeeee;
        }

        /* ================= ORDER ================= */

        .orderInfo {
          margin-top: 12px;
          padding: 11px 14px;
          background: #fbfaf8;
          border: 1px solid #e0ddd4;
          font-size: 12px;
          color: #6b6459;
        }

        .orderInfo strong {
          color: #1f3a5f;
          margin-left: 5px;
        }

        /* ================= MERGE BUTTON ================= */

        .mergeBtn {
          width: 100%;
          margin-top: 16px;
          padding: 13px;
          border: none;
          background: #1f3a5f;
          color: #ffffff;
          border-radius: 4px;
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
        }

        .mergeBtn:hover:not(:disabled) {
          background: #172d49;
        }

        .mergeBtn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
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

        /* ================= DETAILS ================= */

        .detailsBox {
          margin-top: 20px;
          border: 1px solid #e0ddd4;
          background: #fbfaf8;
        }

        .detailsTitle {
          padding: 14px 18px;
          border-bottom: 1px solid #e0ddd4;
          font-size: 11px;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #96723a;
          font-weight: 600;
        }

        .detailRow {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 11px 18px;
          border-bottom: 1px solid #ede9e0;
          font-size: 12px;
        }

        .detailRow:last-child {
          border-bottom: none;
        }

        .detailLeft {
          display: flex;
          align-items: center;
          gap: 9px;
          min-width: 0;
        }

        .detailLeft > span:last-child {
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }

        .detailNumber {
          width: 22px;
          height: 22px;
          border: 1px solid #d7d2c6;
          display: flex;
          align-items: center;
          justify-content: center;
          color: #1f3a5f;
          font-size: 10px;
          flex-shrink: 0;
        }

        .detailRow > strong {
          font-size: 11px;
          color: #1f3a5f;
          white-space: nowrap;
        }

        /* ================= OUTPUT ================= */

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

        .outputCount {
          margin: 20px 0 6px;
          color: #4a4740;
          font-size: 13px;
        }

        .outputInfo {
          font-size: 11px;
          color: #8a8378;
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
          .filesSection,
          .resultsSection {
            margin-left: 22px;
            margin-right: 22px;
          }

          .statsGrid {
            grid-template-columns: 1fr;
          }

          .fileItem {
            flex-wrap: wrap;
          }

          .fileDetails {
            width: calc(100% - 100px);
          }

          .fileActions {
            margin-left: auto;
          }

          .resultHeader {
            align-items: flex-start;
            flex-direction: column;
            gap: 12px;
          }

          .resetBtn {
            width: 100%;
          }

          .detailRow {
            gap: 10px;
          }

        }

      `}</style>

        </div>
    );
}
