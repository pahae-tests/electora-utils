import { useState, useMemo, useEffect } from "react";

const API_URL = "https://electora.emaily.live/api/electeurs";
const PAGE_SIZE_API = 100;
const PAGE_SIZE_DISPLAY = 10;

export default function FetchPage() {
    const [allRows, setAllRows] = useState([]);
    const [search, setSearch] = useState("");
    const [selectedField, setSelectedField] = useState("__ALL__");
    const [displayPage, setDisplayPage] = useState(1);
    const [loading, setLoading] = useState(false);
    const [loadingProgress, setLoadingProgress] = useState("");
    const [error, setError] = useState("");
    const [selectedParrain, setSelectedParrain] = useState("");
    const [parrainSearch, setParrainSearch] = useState("");
    const [showParrainDropdown, setShowParrainDropdown] = useState(false);

    // =========================================================
    // NORMALISATION
    // =========================================================
    const normaliser = (value) => {
        return String(value ?? "")
            .trim()
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "");
    };

    // =========================================================
    // RÉCUPÉRER TOUTES LES PAGES DE L'API
    // =========================================================
    const fetchAllElecteurs = async () => {
        setLoading(true);
        setError("");
        setLoadingProgress("Connexion à l'API...");

        try {
            let page = 1;
            let allData = [];

            while (true) {
                setLoadingProgress(`Récupération de la page ${page}...`);
                const url =
                    `${API_URL}?` +
                    `statut=tous` +
                    `&section=toutes` +
                    `&region=toutes` +
                    `&page=${page}` +
                    `&pageSize=${PAGE_SIZE_API}`;

                const response = await fetch(url);
                if (!response.ok) {
                    throw new Error(`Erreur HTTP ${response.status}`);
                }

                const result = await response.json();
                if (!result || !Array.isArray(result.data)) {
                    throw new Error("La réponse de l'API est invalide.");
                }

                const pageData = result.data;
                allData = [...allData, ...pageData];
                setLoadingProgress(
                    `${allData.length.toLocaleString("fr-FR")} électeurs récupérés...`
                );

                if (pageData.length < PAGE_SIZE_API) {
                    break;
                }
                page++;
            }

            // Suppression des doublons par ID
            const uniqueMap = new Map();
            allData.forEach((electeur) => {
                if (electeur?.id) {
                    uniqueMap.set(electeur.id, electeur);
                } else {
                    uniqueMap.set(JSON.stringify(electeur), electeur);
                }
            });

            const uniqueData = Array.from(uniqueMap.values());
            setAllRows(uniqueData);
            setDisplayPage(1);
            setLoadingProgress("");
        } catch (err) {
            console.error(err);
            setError(err.message || "Impossible de récupérer les électeurs.");
            setAllRows([]);
        } finally {
            setLoading(false);
        }
    };

    // =========================================================
    // CHARGEMENT INITIAL
    // =========================================================
    useEffect(() => {
        fetchAllElecteurs();
    }, []);

    // =========================================================
    // TOUS LES CHAMPS DISPONIBLES
    // =========================================================
    const fields = useMemo(() => {
        const fieldSet = new Set();
        allRows.forEach((row) => {
            Object.keys(row || {}).forEach((key) => {
                fieldSet.add(key);
            });
        });
        return Array.from(fieldSet);
    }, [allRows]);

    // =========================================================
    // LISTE DES PARRAINS UNIQUES
    // =========================================================
    const parrains = useMemo(() => {
        const parrainSet = new Set();
        allRows.forEach((row) => {
            if (row.parrainNom) {
                parrainSet.add(row.parrainNom);
            }
        });
        return Array.from(parrainSet).sort((a, b) => a.localeCompare(b));
    }, [allRows]);

    // =========================================================
    // FILTRAGE DES PARRAINS POUR LA RECHERCHE
    // =========================================================
    const filteredParrains = useMemo(() => {
        const query = normaliser(parrainSearch);
        if (!query) return parrains;
        return parrains.filter((parrain) => normaliser(parrain).includes(query));
    }, [parrains, parrainSearch]);

    // =========================================================
    // FILTRAGE PAR PARRAIN
    // =========================================================
    const filteredByParrain = useMemo(() => {
        if (!selectedParrain) return allRows;
        return allRows.filter((row) => row.parrainNom === selectedParrain);
    }, [allRows, selectedParrain]);

    // =========================================================
    // STATISTIQUES PARRAIN
    // =========================================================
    const parrainStats = useMemo(() => {
        if (!selectedParrain) return null;

        const totalElecteurs = filteredByParrain.length;
        const bureauNonReconnu = filteredByParrain.filter(
            (row) =>
                !row.bureauVote ||
                row.bureauVote.toLowerCase().includes("aucun") ||
                row.bureauVote.toLowerCase().includes("غير معروف")
        ).length;
        const bureauReconnu = totalElecteurs - bureauNonReconnu;

        // Filtrer les électeurs par "ajouté par"
        const motla9aElecteurs = filteredByParrain.filter(
            (row) => row.ajoutePar?.trim().toLowerCase() === "khalid touzani"
        );
        const serviceElecteurs = filteredByParrain.filter(
            (row) => row.ajoutePar?.trim().toLowerCase() !== "khalid touzani"
        );

        // Statistiques pour "motla9a" (ajouté par Khalid Touzani)
        const motla9aReconnu = motla9aElecteurs.filter(
            (row) =>
                row.bureauVote &&
                !row.bureauVote.toLowerCase().includes("aucun") &&
                !row.bureauVote.toLowerCase().includes("غير معروف")
        ).length;
        const motla9aNonReconnu = motla9aElecteurs.length - motla9aReconnu;

        // Statistiques pour "service" (autres)
        const serviceReconnu = serviceElecteurs.filter(
            (row) =>
                row.bureauVote &&
                !row.bureauVote.toLowerCase().includes("aucun") &&
                !row.bureauVote.toLowerCase().includes("غير معروف")
        ).length;
        const serviceNonReconnu = serviceElecteurs.length - serviceReconnu;

        return {
            totalElecteurs,
            bureauNonReconnu,
            bureauReconnu,
            motla9aTotal: motla9aElecteurs.length,
            motla9aReconnu,
            motla9aNonReconnu,
            serviceTotal: serviceElecteurs.length,
            serviceReconnu,
            serviceNonReconnu,
        };
    }, [filteredByParrain, selectedParrain]);

    // =========================================================
    // RECHERCHE SUR TOUTE LA BASE
    // =========================================================
    const filteredRows = useMemo(() => {
        const query = normaliser(search);
        if (!query) return filteredByParrain;

        return filteredByParrain.filter((row) => {
            if (selectedField === "__ALL__") {
                return Object.values(row || {}).some(
                    (value) => normaliser(value).includes(query)
                );
            }
            return normaliser(row?.[selectedField]).includes(query);
        });
    }, [filteredByParrain, search, selectedField]);

    // =========================================================
    // PAGINATION D'AFFICHAGE
    // =========================================================
    const totalDisplayPages = Math.ceil(filteredRows.length / PAGE_SIZE_DISPLAY);
    const paginatedRows = useMemo(() => {
        const start = (displayPage - 1) * PAGE_SIZE_DISPLAY;
        const end = start + PAGE_SIZE_DISPLAY;
        return filteredRows.slice(start, end);
    }, [filteredRows, displayPage]);

    // =========================================================
    // CHANGEMENT DE RECHERCHE
    // =========================================================
    const handleSearchChange = (value) => {
        setSearch(value);
        setDisplayPage(1);
    };

    // =========================================================
    // CHANGEMENT DE CHAMP
    // =========================================================
    const handleFieldChange = (value) => {
        setSelectedField(value);
        setDisplayPage(1);
    };

    // =========================================================
    // CHANGEMENT DE PARRAIN
    // =========================================================
    const handleParrainChange = (value) => {
        setSelectedParrain(value);
        setDisplayPage(1);
        setShowParrainDropdown(false);
    };

    // =========================================================
    // FORMATAGE
    // =========================================================
    const formatValue = (value) => {
        if (value === null || value === undefined) return "";
        if (typeof value === "boolean") return value ? "Oui" : "Non";
        if (typeof value === "object") {
            try {
                return JSON.stringify(value);
            } catch {
                return String(value);
            }
        }
        return String(value);
    };

    // =========================================================
    // RESET RECHERCHE
    // =========================================================
    const resetSearch = () => {
        setSearch("");
        setDisplayPage(1);
    };

    // =========================================================
    // RESET FILTRE PARRAIN
    // =========================================================
    const resetParrainFilter = () => {
        setSelectedParrain("");
        setParrainSearch("");
        setDisplayPage(1);
    };

    // =========================================================
    // PAGINATION
    // =========================================================
    const previousPage = () => {
        setDisplayPage((current) => Math.max(1, current - 1));
    };

    const nextPage = () => {
        setDisplayPage((current) => Math.min(totalDisplayPages, current + 1));
    };

    // =========================================================
    // AFFICHAGE
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
                        <span className="eyebrow">Base électorale</span>
                        <h1>Recherche des électeurs</h1>
                        <p>
                            Recherchez directement dans l'ensemble des électeurs disponibles via l'API.
                        </p>
                    </div>
                </header>

                {/* =================================================
                    API
                ================================================= */}
                <section className="apiSection">
                    <div className="apiStatus">
                        <span
                            className={
                                loading
                                    ? "statusDot loadingDot"
                                    : error
                                    ? "statusDot errorDot"
                                    : "statusDot"
                            }
                        />
                        <div>
                            <span className="statusTitle">
                                {loading
                                    ? "Récupération des données..."
                                    : error
                                    ? "Erreur de connexion"
                                    : "API connectée"}
                            </span>
                            <span className="apiUrl">{API_URL}</span>
                        </div>
                    </div>
                    <button
                        type="button"
                        className="refreshBtn"
                        onClick={fetchAllElecteurs}
                        disabled={loading}
                    >
                        {loading ? "Chargement..." : "Actualiser"}
                    </button>
                </section>

                {/* =================================================
                    PROGRESSION
                ================================================= */}
                {loading && (
                    <div className="progressBox">
                        <div className="loader" />
                        <span>{loadingProgress}</span>
                    </div>
                )}

                {/* =================================================
                    ERROR
                ================================================= */}
                {error && <div className="error">{error}</div>}

                {/* =================================================
                    FILTRE PARRAIN
                ================================================= */}
                {!loading && allRows.length > 0 && (
                    <section className="parrainFilterSection">
                        <div className="parrainFilterHeader">
                            <div>
                                <span className="sectionLabel">Filtre par parrain</span>
                                <h2>Sélectionner un parrain</h2>
                            </div>
                        </div>
                        <div className="parrainSearchWrapper">
                            <div className="searchInputWrapper parrainSearchInput">
                                <span className="searchIcon">⌕</span>
                                <input
                                    type="text"
                                    value={parrainSearch}
                                    onChange={(e) => setParrainSearch(e.target.value)}
                                    onFocus={() => setShowParrainDropdown(true)}
                                    placeholder="Rechercher un parrain..."
                                />
                                {parrainSearch && (
                                    <button
                                        type="button"
                                        className="clearSearch"
                                        onClick={() => {
                                            setParrainSearch("");
                                            setShowParrainDropdown(true);
                                        }}
                                    >
                                        ×
                                    </button>
                                )}
                            </div>
                            {showParrainDropdown && (
                                <div className="parrainDropdown">
                                    {filteredParrains.length > 0 ? (
                                        filteredParrains.map((parrain) => (
                                            <div
                                                key={parrain}
                                                className="parrainOption"
                                                onClick={() => handleParrainChange(parrain)}
                                            >
                                                {parrain}
                                            </div>
                                        ))
                                    ) : (
                                        <div className="parrainOption noResults">
                                            Aucun parrain trouvé
                                        </div>
                                    )}
                                </div>
                            )}
                        </div>
                        {selectedParrain && parrainStats && (
                            <div className="parrainStats">
                                <div className="statGroup">
                                    <div className="statItem">
                                        <span>Total électeurs:</span>
                                        <strong>{parrainStats.totalElecteurs.toLocaleString("fr-FR")}</strong>
                                    </div>
                                    <div className="statItem">
                                        <span>Bureau non reconnu:</span>
                                        <strong>{parrainStats.bureauNonReconnu.toLocaleString("fr-FR")}</strong>
                                    </div>
                                    <div className="statItem">
                                        <span>Bureau reconnu:</span>
                                        <strong>{parrainStats.bureauReconnu.toLocaleString("fr-FR")}</strong>
                                    </div>
                                </div>
                                <div className="statGroup">
                                    <h3>Ajouté par : Khalid Touzani (motla9a)</h3>
                                    <div className="statItem">
                                        <span>Total:</span>
                                        <strong>{parrainStats.motla9aTotal.toLocaleString("fr-FR")}</strong>
                                    </div>
                                    <div className="statItem">
                                        <span>Reconnu:</span>
                                        <strong>{parrainStats.motla9aReconnu.toLocaleString("fr-FR")}</strong>
                                    </div>
                                    <div className="statItem">
                                        <span>Non reconnu:</span>
                                        <strong>{parrainStats.motla9aNonReconnu.toLocaleString("fr-FR")}</strong>
                                    </div>
                                </div>
                                <div className="statGroup">
                                    <h3>Ajouté par : Autres (service)</h3>
                                    <div className="statItem">
                                        <span>Total:</span>
                                        <strong>{parrainStats.serviceTotal.toLocaleString("fr-FR")}</strong>
                                    </div>
                                    <div className="statItem">
                                        <span>Reconnu:</span>
                                        <strong>{parrainStats.serviceReconnu.toLocaleString("fr-FR")}</strong>
                                    </div>
                                    <div className="statItem">
                                        <span>Non reconnu:</span>
                                        <strong>{parrainStats.serviceNonReconnu.toLocaleString("fr-FR")}</strong>
                                    </div>
                                </div>
                            </div>
                        )}
                        {selectedParrain && (
                            <button
                                type="button"
                                className="secondaryBtn resetParrainBtn"
                                onClick={resetParrainFilter}
                            >
                                Réinitialiser le filtre parrain
                            </button>
                        )}
                    </section>
                )}

                {/* =================================================
                    SEARCH
                ================================================= */}
                {!loading && allRows.length > 0 && (
                    <section className="searchSection">
                        <div className="searchHeader">
                            <div>
                                <span className="sectionLabel">Recherche globale</span>
                                <h2>Trouver un électeur</h2>
                            </div>
                            <div className="stats">
                                <span>{filteredRows.length.toLocaleString("fr-FR")}</span>
                                <small>
                                    {selectedParrain
                                        ? `électeurs pour "${selectedParrain}"`
                                        : "électeurs chargés"}
                                </small>
                            </div>
                        </div>
                        <div className="searchGrid">
                            <div className="field">
                                <label htmlFor="field">Rechercher dans</label>
                                <select
                                    id="field"
                                    value={selectedField}
                                    onChange={(e) => handleFieldChange(e.target.value)}
                                >
                                    <option value="__ALL__">Tous les champs</option>
                                    {fields.map((field) => (
                                        <option key={field} value={field}>
                                            {field}
                                        </option>
                                    ))}
                                </select>
                            </div>
                            <div className="field">
                                <label htmlFor="search">Terme de recherche</label>
                                <div className="searchInputWrapper">
                                    <span className="searchIcon">⌕</span>
                                    <input
                                        id="search"
                                        type="text"
                                        value={search}
                                        onChange={(e) => handleSearchChange(e.target.value)}
                                        placeholder={
                                            selectedField === "__ALL__"
                                                ? "Nom, CIN, téléphone, bureau..."
                                                : `Rechercher dans « ${selectedField} »`
                                        }
                                    />
                                    {search && (
                                        <button
                                            type="button"
                                            className="clearSearch"
                                            onClick={resetSearch}
                                        >
                                            ×
                                        </button>
                                    )}
                                </div>
                            </div>
                        </div>
                        <div className="searchInfo">
                            <span>
                                {search ? (
                                    <>
                                        Recherche globale pour <strong>« {search} »</strong>
                                    </>
                                ) : selectedParrain ? (
                                    `Filtre appliqué : Parrain = "${selectedParrain}"`
                                ) : (
                                    "Recherche dans l'ensemble des électeurs récupérés"
                                )}
                            </span>
                            <strong className="resultCount">
                                {filteredRows.length.toLocaleString("fr-FR")} résultat
                                {filteredRows.length !== 1 ? "s" : ""}
                            </strong>
                        </div>
                    </section>
                )}

                {/* =================================================
                    RESULTS
                ================================================= */}
                {!loading && allRows.length > 0 && (
                    <section className="results">
                        <div className="resultsHeader">
                            <div>
                                <h2>Registre des électeurs</h2>
                                <p>
                                    {search
                                        ? `${filteredRows.length.toLocaleString("fr-FR")} résultat(s) trouvé(s)`
                                        : selectedParrain
                                        ? `${filteredRows.length.toLocaleString("fr-FR")} électeurs pour "${selectedParrain}"`
                                        : `${allRows.length.toLocaleString("fr-FR")} électeurs récupérés`}
                                </p>
                            </div>
                            {(search || selectedParrain) && (
                                <button
                                    type="button"
                                    className="secondaryBtn"
                                    onClick={() => {
                                        resetSearch();
                                        resetParrainFilter();
                                    }}
                                >
                                    Réinitialiser tout
                                </button>
                            )}
                        </div>

                        {/* TABLE */}
                        {paginatedRows.length > 0 ? (
                            <div className="tableWrap">
                                <table>
                                    <thead>
                                        <tr>
                                            <th className="numberColumn">#</th>
                                            <th>CIN</th>
                                            <th>Parrain</th>
                                            <th>Nom</th>
                                            <th>Prénom</th>
                                            <th>Nom arabe</th>
                                            <th>Prénom arabe</th>
                                            <th>Téléphone</th>
                                            <th>Bureau de vote</th>
                                            <th>Numéro de vote</th>
                                            <th>Ajouté par</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {paginatedRows.map((row, rowIndex) => {
                                            const realIndex =
                                                (displayPage - 1) * PAGE_SIZE_DISPLAY +
                                                rowIndex;
                                            return (
                                                <tr key={row.id || realIndex}>
                                                    <td className="rowNumber">{realIndex + 1}</td>
                                                    <td>{row.numeroCIN || ""}</td>
                                                    <td>{row.parrainNom || ""}</td>
                                                    <td>{row.nom || ""}</td>
                                                    <td>{row.prenom || ""}</td>
                                                    <td dir="rtl">{row.nomArabe || ""}</td>
                                                    <td dir="rtl">{row.prenomArabe || ""}</td>
                                                    <td>{row.telephone || ""}</td>
                                                    <td>{row.bureauVote || ""}</td>
                                                    <td>{row.numeroVote || ""}</td>
                                                    <td>{row.ajoutePar || ""}</td>
                                                </tr>
                                            );
                                        })}
                                    </tbody>
                                </table>
                            </div>
                        ) : (
                            <div className="noResults">
                                <div className="noResultsIcon">—</div>
                                <div className="noResultsTitle">Aucun résultat</div>
                                <div className="noResultsText">
                                    Aucun électeur ne correspond à votre recherche.
                                </div>
                            </div>
                        )}

                        {/* PAGINATION */}
                        {filteredRows.length > 0 && (
                            <div className="pagination">
                                <button
                                    type="button"
                                    className="paginationBtn"
                                    onClick={previousPage}
                                    disabled={displayPage <= 1}
                                >
                                    ← Précédent
                                </button>
                                <div className="pageNumber">
                                    <span>Page</span>
                                    <strong>{displayPage}</strong>
                                    <span>sur</span>
                                    <strong>{totalDisplayPages}</strong>
                                </div>
                                <button
                                    type="button"
                                    className="paginationBtn"
                                    onClick={nextPage}
                                    disabled={displayPage >= totalDisplayPages}
                                >
                                    Suivant →
                                </button>
                            </div>
                        )}

                        {/* TOTAL */}
                        <div className="totalRow">
                            <span>
                                {selectedParrain
                                    ? `Électeurs pour "${selectedParrain}"`
                                    : "Électeurs correspondant à la recherche"}
                            </span>
                            <strong>{filteredRows.length.toLocaleString("fr-FR")}</strong>
                        </div>
                    </section>
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
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                    color: #1e2124;
                }

                .sheet {
                    width: 100%;
                    max-width: 1500px;
                    background: #ffffff;
                    border: 1px solid #e0ddd4;
                    border-radius: 4px;
                    overflow: hidden;
                }

                /* HEADER */
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

                /* API */
                .apiSection {
                    margin: 26px 36px 0;
                    padding: 14px 16px;
                    background: #fbfaf8;
                    border: 1px solid #e0ddd4;
                    border-radius: 4px;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    gap: 20px;
                }

                .apiStatus {
                    display: flex;
                    align-items: center;
                    gap: 11px;
                    min-width: 0;
                }

                .statusDot {
                    width: 8px;
                    height: 8px;
                    border-radius: 50%;
                    background: #3f7657;
                    flex-shrink: 0;
                }

                .loadingDot {
                    background: #96723a;
                }

                .errorDot {
                    background: #8a2d2d;
                }

                .apiStatus > div {
                    display: flex;
                    flex-direction: column;
                    gap: 2px;
                    min-width: 0;
                }

                .statusTitle {
                    font-size: 12px;
                    font-weight: 600;
                    color: #4a4740;
                }

                .apiUrl {
                    color: #8a8378;
                    font-size: 11px;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    white-space: nowrap;
                }

                .refreshBtn {
                    background: transparent;
                    border: 1px solid #1f3a5f;
                    color: #1f3a5f;
                    border-radius: 4px;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                    cursor: pointer;
                    flex-shrink: 0;
                }

                .refreshBtn:hover:not(:disabled) {
                    background: #eef2f6;
                }

                .refreshBtn:disabled {
                    opacity: 0.45;
                    cursor: not-allowed;
                }

                /* PROGRESS */
                .progressBox {
                    margin: 20px 36px;
                    padding: 18px;
                    border: 1px solid #e0ddd4;
                    background: #fbfaf8;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 12px;
                    color: #6b6459;
                    font-size: 13px;
                }

                .loader {
                    width: 19px;
                    height: 19px;
                    border: 2px solid #d7d2c6;
                    border-top-color: #1f3a5f;
                    border-radius: 50%;
                    animation: spin 0.8s linear infinite;
                }

                @keyframes spin {
                    to {
                        transform: rotate(360deg);
                    }
                }

                /* ERROR */
                .error {
                    margin: 20px 36px 4px;
                    background: #fbeeee;
                    border: 1px solid #e3b8b8;
                    color: #8a2d2d;
                    border-radius: 4px;
                    padding: 10px 14px;
                    font-size: 13px;
                }

                /* PARRAIN FILTER */
                .parrainFilterSection {
                    margin: 28px 36px 0;
                    padding: 24px 0;
                    border-top: 1px solid #e0ddd4;
                    border-bottom: 1px solid #e0ddd4;
                }

                .parrainFilterHeader {
                    display: flex;
                    justify-content: space-between;
                    align-items: flex-start;
                    margin-bottom: 18px;
                }

                .parrainSearchWrapper {
                    position: relative;
                    margin-bottom: 18px;
                }

                .parrainSearchInput {
                    width: 100%;
                }

                .parrainDropdown {
                    position: absolute;
                    top: 100%;
                    left: 0;
                    right: 0;
                    max-height: 200px;
                    overflow-y: auto;
                    background: #ffffff;
                    border: 1px solid #e0ddd4;
                    border-radius: 4px;
                    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                    z-index: 1000;
                    margin-top: 4px;
                }

                .parrainOption {
                    padding: 10px 12px;
                    cursor: pointer;
                    font-size: 14px;
                    color: #1e2124;
                }

                .parrainOption:hover {
                    background: #fbfaf8;
                }

                .parrainOption.noResults {
                    color: #8a8378;
                    font-style: italic;
                }

                .parrainStats {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                    gap: 20px;
                    padding: 14px;
                    background: #fbfaf8;
                    border: 1px solid #e0ddd4;
                    border-radius: 4px;
                }

                .statGroup {
                    display: flex;
                    flex-direction: column;
                    gap: 8px;
                }

                .statGroup h3 {
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 14px;
                    color: #1f3a5f;
                    margin: 0 0 8px 0;
                    border-bottom: 1px solid #d7d2c6;
                    padding-bottom: 4px;
                }

                .statItem {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    font-size: 13px;
                    color: #4a4740;
                }

                .statItem strong {
                    color: #1f3a5f;
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 16px;
                }

                .resetParrainBtn {
                    margin-top: 14px;
                }

                /* SEARCH */
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
                    font-family: Georgia, "Times New Roman", serif;
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
                    font-family: Georgia, "Times New Roman", serif;
                    color: #1f3a5f;
                    font-size: 22px;
                    font-weight: 600;
                }

                .stats small {
                    font-size: 12px;
                }

                .searchGrid {
                    display: grid;
                    grid-template-columns: 280px 1fr;
                    gap: 18px;
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
                    box-shadow: 0 0 0 3px rgba(31, 58, 95, 0.12);
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

                /* RESULTS */
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
                    font-family: Georgia, "Times New Roman", serif;
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

                /* TABLE */
                .tableWrap {
                    overflow: auto;
                    max-height: 650px;
                    border: 1px solid #e0ddd4;
                    border-radius: 4px;
                }

                table {
                    width: 100%;
                    min-width: 1200px;
                    border-collapse: collapse;
                    font-size: 12px;
                }

                th {
                    position: sticky;
                    top: 0;
                    z-index: 2;
                    text-align: left;
                    padding: 10px 12px;
                    background: #fbfaf8;
                    color: #6b6459;
                    font-weight: 600;
                    font-size: 11px;
                    border-bottom: 2px solid #1f3a5f;
                    white-space: nowrap;
                }

                td {
                    text-align: left;
                    padding: 9px 12px;
                    border-bottom: 1px solid #ede9e0;
                    white-space: nowrap;
                    color: #1e2124;
                    max-width: 300px;
                    overflow: hidden;
                    text-overflow: ellipsis;
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
                    width: 45px;
                    text-align: center;
                }

                .rowNumber {
                    color: #9a9488;
                    text-align: center;
                    font-size: 11px;
                }

                /* NO RESULTS */
                .noResults {
                    padding: 55px 20px;
                    border: 1px solid #e0ddd4;
                    background: #fbfaf8;
                    text-align: center;
                }

                .noResultsIcon {
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 30px;
                    color: #96723a;
                    margin-bottom: 10px;
                }

                .noResultsTitle {
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 17px;
                    color: #16191c;
                    margin-bottom: 6px;
                }

                .noResultsText {
                    color: #8a8378;
                    font-size: 13px;
                }

                /* PAGINATION */
                .pagination {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-top: 18px;
                    padding-top: 16px;
                    border-top: 1px solid #e0ddd4;
                }

                .paginationBtn {
                    background: transparent;
                    border: 1px solid #1f3a5f;
                    color: #1f3a5f;
                    border-radius: 4px;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                    cursor: pointer;
                }

                .paginationBtn:hover:not(:disabled) {
                    background: #eef2f6;
                }

                .paginationBtn:disabled {
                    opacity: 0.35;
                    cursor: not-allowed;
                }

                .pageNumber {
                    display: flex;
                    align-items: center;
                    gap: 7px;
                    color: #8a8378;
                    font-size: 12px;
                }

                .pageNumber strong {
                    color: #1f3a5f;
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 16px;
                }

                /* TOTAL */
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

                /* RESPONSIVE */
                @media (max-width: 700px) {
                    .page {
                        padding: 20px 10px;
                    }

                    .letterheadText {
                        padding: 25px 22px 22px;
                    }

                    .apiSection,
                    .parrainFilterSection,
                    .searchSection,
                    .results,
                    .progressBox {
                        margin-left: 22px;
                        margin-right: 22px;
                    }

                    .searchGrid {
                        grid-template-columns: 1fr;
                    }

                    .apiSection {
                        align-items: flex-start;
                        flex-direction: column;
                    }

                    .refreshBtn {
                        width: 100%;
                    }

                    .searchInfo {
                        flex-direction: column;
                        align-items: flex-start;
                        gap: 6px;
                    }

                    .parrainStats {
                        grid-template-columns: 1fr;
                    }
                }
            `}</style>
        </div>
    );
}