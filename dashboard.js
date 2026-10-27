(() => {

    "use strict";


    const SHEETS_API =
        "https://sheets.googleapis.com/v4/spreadsheets";

    const DRIVE_API =
        "https://www.googleapis.com/drive/v3/files";


    const SESSION_KEY = "fmAssetSession";
    let tokenRequestPromise = null;
    const auditOpen = new Map();
    const auditLimits = new Map();
    const auditOlderLimits = new Map();
    const AUDIT_PAGE_SIZE = 20;
    const auditDateFilter = { mode: "today", day: "", from: "", to: "" };
    const AUDIT_UK_FORMATTER = new Intl.DateTimeFormat("en-GB", {
        timeZone: "Europe/London", year: "numeric", month: "2-digit", day: "2-digit"
    });
    const SHEIN_ESTIMATE_PREF = "fmSheinOutboundEstimateEnabled";
    let sheinEstimateEnabled = true;

    // Each graph opens on a rolling week. Ranges are deliberately independent.
    const graphRanges = {client: "7d", movement: "7d", quantity: "7d"};
    let currentClientGraph = null;
    let currentWarehouseGraph = null;


    let movementSubmitBusy = false;
    let batchForm = null;

    const state = {

        idTokenPayload: null,

        accessToken: null,

        workbookId: null,

        workbookName: null,

        inventory: [],

        transactions: [],
        assetAvailability: [],
        assetOperationRows: [],
        clientAdjustmentRows: [],

        // The live inventory/transactions above remain the source for
        // movement recording and inventory management. Dashboard-only
        // views can switch to an archived snapshot without changing them.
        dashboardInventory: [],
        dashboardTransactions: [],
        dashboardClients: [],
        historicalLedgerFiles: [],
        dashboardSource: { kind: "current", id: null, name: "Current live ledger" },

        todayLoads: [],
        routines: [],
        alerts: [],
        alertHeaders: [],

        clients: [],
        brokers: [...(window.FM_BROKERS?.DEFAULT_BROKERS || [])],
        clientListRows: [],
        clientTypes: { categories: [], lookup: Object.create(null) },
        auditVisibleRows: [],
        clientDetailsRows: [],
        clientDetailsVisibleCount: 5,
        globalSearchRows: []
    };


    const $ =
        id => document.getElementById(id);

document.addEventListener(
    "DOMContentLoaded",
    () => {

        batchForm = FM_BATCH_FORM?.create?.({
            containerId: "client-movement-entries",
            addButtonId: "add-client-entry",
            selfClient: CONFIG.SELF_CLIENT_NAME || "HSC London (Self)",
            getClients: () => state.clients,
            getBrokers: () => state.brokers,
            getAssets: () => state.inventory.map(item => item.asset),
            onChange: () => updateMovementPreview()
        });

        bindEvents();

        bindQuickMenu();

        setDefaultTimestamp();

        handleMovementChange();

        waitForGoogle();

    }
);


    function bindEvents() {
        // Event delegation survives graph redraws and keeps the range controls usable on touch devices.
        document.addEventListener("click", event => {
            const control = event.target.closest?.("button[data-graph-range]");
            if (!control) return;
            const group = control.closest("[data-graph-range-group]")?.dataset.graphRangeGroup;
            const range = control.dataset.graphRange;
            if (!Object.prototype.hasOwnProperty.call(graphRanges, group) || !["7d", "30d", "90d", "all"].includes(range)) return;
            graphRanges[group] = range;
            syncGraphRangeControls(group);
            if (group === "client" && currentClientGraph) {
                renderClientAssetGraph(currentClientGraph.client, currentClientGraph.asset);
            } else if (group === "movement" && currentWarehouseGraph) {
                renderWarehouseAssetMovementGraph(currentWarehouseGraph.asset, currentWarehouseGraph.transactions);
            } else if (group === "quantity" && currentWarehouseGraph) {
                renderWarehouseAssetQuantityGraph(currentWarehouseGraph.asset, currentWarehouseGraph.transactions, currentWarehouseGraph.balance);
            }
        });


        // Always call from a direct user gesture with silent=false. Passing the
        // click Event as the argument makes a truthy "silent" flag and prevents
        // the first-time Google consent window from appearing.
        $("grant-access").addEventListener("click", () => requestSheetAccess(false));


        $("refresh")
            .addEventListener(
                "click",
                loadDashboard
            );


        $("sign-out")
            .addEventListener(
                "click",
                signOut
            );


        $("movement-form")
            .addEventListener(
                "submit",
                recordMovement
            );


        $("movement-photo")?.addEventListener("change", handleDashboardPhotoChange);
        $("clear-photo")?.addEventListener("click", clearDashboardPhoto);
        $("take-photo")?.addEventListener("click", () => $("movement-photo-camera")?.click());
        $("movement-photo-camera")?.addEventListener("change", event => {
            const file = event.target.files?.[0];
            const main = $("movement-photo");
            if (!file || !main) return;
            try {
                const transfer = new DataTransfer();
                transfer.items.add(file);
                main.files = transfer.files;
            } catch (_) {
                // Older mobile browsers may not expose DataTransfer for file inputs.
                window._fmCameraFile = file;
            }
            handleDashboardPhotoChange();
        });


        $("client-cards")
            .addEventListener(
                "click",
                event => {
                    const card = event.target.closest("[data-client]");
                    if (card) {
                        openClientDetails(card.dataset.client);
                    }
                }
            );

        $("client-cards")
            .addEventListener(
                "keydown",
                event => {
                    if (event.key !== "Enter" && event.key !== " ") return;
                    const card = event.target.closest("[data-client]");
                    if (!card) return;
                    event.preventDefault();
                    openClientDetails(card.dataset.client);
                }
            );

        $("client-search")
            .addEventListener(
                "input",
                renderClientCards
            );

        $("global-search")?.addEventListener("input", renderGlobalSearch);
        $("clear-global-search")?.addEventListener("click", () => {
            if ($("global-search")) $("global-search").value = "";
            renderGlobalSearch();
            $("global-search")?.focus();
        });
        $("global-search-results")?.addEventListener("click", event => {
            const info = event.target.closest("[data-global-info-index]");
            if (info) {
                const group = state.globalSearchRows[Number(info.dataset.globalInfoIndex)];
                if (group) {
                    prepareInfoModal(group);
                    $("info-modal").classList.remove("hidden");
                    requestAnimationFrame(() => loadInfoPhoto(group));
                }
                return;
            }
            const client = event.target.closest("[data-global-client]");
            if (client) { openClientDetails(client.dataset.globalClient); return; }
            const asset = event.target.closest("[data-global-asset]");
            if (asset) openWarehouseAssetDetails(asset.dataset.globalAsset);
        });

        const dashboardFilterPairs = [
            ["dashboard-filter-client", "dashboard-filter-client-audit"],
            ["dashboard-filter-asset", "dashboard-filter-asset-audit"]
        ];

        dashboardFilterPairs.forEach(([mainId, auditId]) => {
            $(mainId)?.addEventListener("change", () => {
                if ($(auditId)) $(auditId).value = $(mainId).value;
                resetAuditPaging();
                renderClientCards();
                renderWarehouseAssets();
                renderAudit();
                renderFilteredLoadTables();
            });
            $(auditId)?.addEventListener("change", () => {
                if ($(mainId)) $(mainId).value = $(auditId).value;
                resetAuditPaging();
                renderClientCards();
                renderWarehouseAssets();
                renderAudit();
                renderFilteredLoadTables();
            });
        });

        // The dashboard's day filter can select an audit date, but dates chosen
        // within the audit panel do not change the warehouse or load dashboard.
        $("dashboard-filter-date")?.addEventListener("change", () => {
            auditDateFilter.day = $("dashboard-filter-date").value;
            auditDateFilter.mode = auditDateFilter.day ? "day" : "today";
            resetAuditPaging();
            renderClientCards();
            renderWarehouseAssets();
            renderAudit();
            renderFilteredLoadTables();
        });
        $("dashboard-filter-date-audit")?.addEventListener("change", event => {
            auditDateFilter.day = event.target.value;
            auditDateFilter.mode = auditDateFilter.day ? "day" : "today";
            resetAuditPaging();
            renderAudit();
        });
        $("audit-trail")?.addEventListener("click", event => {
            const periodButton = event.target.closest("button[data-audit-period]");
            if (!periodButton) return;
            const period = periodButton.dataset.auditPeriod;
            if (!["today", "yesterday", "7d", "30d", "all", "custom"].includes(period)) return;
            auditDateFilter.mode = period;
            if (period === "custom" && (!auditDateFilter.from || !auditDateFilter.to)) {
                const today = auditUkDayKey(new Date());
                auditDateFilter.from = auditSubtractDays(today, 6);
                auditDateFilter.to = today;
            }
            resetAuditPaging();
            renderAudit();
        });
        ["from", "to"].forEach(boundary => {
            $("audit-range-" + boundary)?.addEventListener("change", event => {
                auditDateFilter[boundary] = event.target.value;
                auditDateFilter.mode = "custom";
                resetAuditPaging();
                renderAudit();
            });
        });

        $("clear-dashboard-filters")?.addEventListener("click", clearDashboardFilters);
        $("clear-audit-filters")?.addEventListener("click", clearDashboardFilters);

        $("dashboard-data-version")?.addEventListener("change", handleDashboardDataVersionChange);

        $("audit-trail")?.addEventListener("click", event => {
            const more = event.target.closest("[data-audit-more]");
            if (more) {
                const key = more.dataset.auditMore;
                auditLimits.set(key, (auditLimits.get(key) || AUDIT_PAGE_SIZE) + AUDIT_PAGE_SIZE);
                renderAudit();
                return;
            }
            const older = event.target.closest("[data-audit-older]");
            if (older) {
                const key = older.dataset.auditOlder;
                auditOlderLimits.set(key, (auditOlderLimits.get(key) || 0) + AUDIT_PAGE_SIZE);
                renderAudit();
                return;
            }
            const button = event.target.closest("[data-info-index]");
            if (button) openInfoModal(Number(button.dataset.infoIndex));
        });
        $("audit-trail")?.addEventListener("toggle", event => {
            const details = event.target;
            if (details.matches?.("details[data-audit-category]")) {
                auditOpen.set(details.dataset.auditCategory, details.open);
            }
        }, true);
        try { sheinEstimateEnabled = localStorage.getItem(SHEIN_ESTIMATE_PREF) !== "false"; }
        catch (_) { sheinEstimateEnabled = true; }
        const estimateToggle = $("shein-outbound-estimate");
        if (estimateToggle) {
            estimateToggle.checked = sheinEstimateEnabled;
            estimateToggle.addEventListener("change", () => {
                sheinEstimateEnabled = estimateToggle.checked;
                try { localStorage.setItem(SHEIN_ESTIMATE_PREF, String(sheinEstimateEnabled)); } catch (_) {}
                renderFilteredLoadTables();
            });
        }
        $("client-details-transactions")?.addEventListener("click", event => {
            const loadMore = event.target.closest("[data-client-load-more]");
            if (loadMore) {
                state.clientDetailsVisibleCount += 5;
                renderClientDetailsTransactionList();
                return;
            }
            const button = event.target.closest("[data-client-info-index]");
            if (button) openClientTransactionInfo(Number(button.dataset.clientInfoIndex));
        });
        $("close-info")?.addEventListener("click", closeInfoModal);
        document.addEventListener("click", event => {
            const summaryToggle = event.target.closest("[data-summary-toggle]");
            if (summaryToggle) {
                toggleTransactionSummary(summaryToggle);
                return;
            }
            if (event.target.matches("[data-close-info]")) closeInfoModal();
        });
        document.addEventListener("click", event => {
            const exportButton = event.target.closest("[data-export-client-ledger]");
            if (exportButton) {
                openClientLedgerExport();
                return;
            }
            if (event.target.matches("[data-close-client-ledger-export]")) closeClientLedgerExport();
        });
        $("close-client-ledger-export")?.addEventListener("click", closeClientLedgerExport);
        $("cancel-client-ledger-export")?.addEventListener("click", closeClientLedgerExport);
        $("run-client-ledger-export")?.addEventListener("click", runClientLedgerExport);
        $("run-client-ledger-pdf")?.addEventListener("click", runClientLedgerPdfExport);
        $("run-client-ledger-week-pdf")?.addEventListener("click", runClientLedgerWeeklyPdfExport);
        $("client-ledger-export-select-all")?.addEventListener("click", selectAllClientLedgerExportAssets);
        $("client-ledger-export-assets")?.addEventListener("change", handleClientLedgerAssetSelectionChange);
        $("clear-client-details-filters")?.addEventListener("click", () => {
            const client = $("client-details-title")?.textContent || "";
            if ($("client-details-filter-date")) $("client-details-filter-date").value = "";
            if ($("client-details-filter-asset")) $("client-details-filter-asset").value = "";
            if (client) openClientDetails(client);
        });

        $("warehouse-assets")?.addEventListener("click", event => {
            const card = event.target.closest("[data-asset-name]");
            if (!card) return;
            openWarehouseAssetDetails(card.dataset.assetName);
        });

        $("warehouse-assets")?.addEventListener("keydown", event => {
            if (event.key !== "Enter" && event.key !== " ") return;
            const card = event.target.closest("[data-asset-name]");
            if (!card) return;
            event.preventDefault();
            openWarehouseAssetDetails(card.dataset.assetName);
        });

        $("close-warehouse-asset")?.addEventListener("click", closeWarehouseAssetDetails);
        document.addEventListener("click", event => {
            if (event.target.matches("[data-close-warehouse-asset]")) closeWarehouseAssetDetails();
        });

$("client-details-summary")
            .addEventListener(
                "click",
                event => {
                    const button = event.target.closest("[data-client-asset]");
                    if (!button) return;

                    const client = $("client-details-title").textContent;
                    renderClientAssetGraph(client, button.dataset.clientAsset);

                    document
                        .querySelectorAll("[data-client-asset]")
                        .forEach(item => item.classList.remove("selected"));

                    button.classList.add("selected");
                }
            );


        $("close-client-details")
            .addEventListener(
                "click",
                closeClientDetails
            );


        document.addEventListener(
            "click",
            event => {

                if (
                    event.target.matches(
                        "[data-close-client-details]"
                    )
                ) {

                    closeClientDetails();

                }

            }
        );


        $("manage-inventory")
            .addEventListener(
                "click",
                openInventoryManager
            );

        $("save-availability")?.addEventListener("click", saveAvailabilityCounts);


        $("close-inventory")
            .addEventListener(
                "click",
                closeInventoryManager
            );


        document.addEventListener(
            "click",
            event => {

                if (
                    event.target.matches(
                        "[data-close-modal]"
                    )
                ) {

                    closeInventoryManager();

                }

            }
        );


        $("add-asset")
            .addEventListener(
                "click",
                addAssetType
            );


        $("manage-inventory-body")
            .addEventListener(
                "click",
                event => {

                    const button =
                        event.target.closest(
                            "button[data-action]"
                        );

                    if (!button) {
                        return;
                    }


                    const rowNumber =
                        Number(button.dataset.row);


                    if (
                        button.dataset.action ===
                        "delete"
                    ) {

                        deleteAssetType(
                            rowNumber
                        );

                    }


                    if (
                        button.dataset.action ===
                        "rename"
                    ) {

                        renameAssetType(
                            rowNumber
                        );

                    }

                }
            );

    }



function bindQuickMenu() {
    const nav = document.querySelector(".quick-nav");
    const toggle = document.getElementById("menu-toggle");
    const links = document.getElementById("quick-links");

    if (!nav || !toggle || !links) return;

    const storageKey = "fmQuickMenuOpen";
    const mobileQuery = window.matchMedia("(max-width: 760px)");
    const isMobile = () => mobileQuery.matches;

    function setOpen(open, persist = true) {
        const mobile = isMobile();
        const next = mobile ? Boolean(open) : Boolean(open);
        nav.classList.toggle("menu-open", next);
        toggle.setAttribute("aria-expanded", String(next));
        if (persist && !mobile) {
            try { localStorage.setItem(storageKey, next ? "1" : "0"); } catch (_) {}
        }
    }

    let saved = null;
    try { saved = localStorage.getItem(storageKey); } catch (_) {}

    // Mobile always starts collapsed. Desktop can remember the user's preferred state.
    setOpen(isMobile() ? false : saved === "1");

    toggle.addEventListener("click", event => {
        event.preventDefault();
        event.stopPropagation();
        setOpen(!nav.classList.contains("menu-open"));
    });

    links.addEventListener("click", event => {
        const link = event.target.closest("a");
        if (!link) return;
        if (isMobile()) setOpen(false, false);
    });

    const handleViewportChange = event => {
        if (event.matches) setOpen(false, false);
    };
    if (mobileQuery.addEventListener) mobileQuery.addEventListener("change", handleViewportChange);
    else if (mobileQuery.addListener) mobileQuery.addListener(handleViewportChange);
}

    function waitForGoogle() {
        if (window.google?.accounts?.id && window.google?.accounts?.oauth2) {
            initializeGoogle();
            return;
        }

        let finished = false;
        const started = Date.now();

        const onLoaded = () => {
            if (finished) return;
            finished = true;
            window.removeEventListener("fm-google-loaded", onLoaded);
            initializeGoogle();
        };

        window.addEventListener("fm-google-loaded", onLoaded, { once: true });

        const timer = setInterval(() => {
            if (finished) {
                clearInterval(timer);
                return;
            }

            if (window.google?.accounts?.id && window.google?.accounts?.oauth2) {
                clearInterval(timer);
                finished = true;
                window.removeEventListener("fm-google-loaded", onLoaded);
                initializeGoogle();
                return;
            }

            if (Date.now() - started >= 8000) {
                clearInterval(timer);
                finished = true;
                window.removeEventListener("fm-google-loaded", onLoaded);
                setAuthStatus("Google services could not be loaded. Check your internet connection.", true);
            }
        }, 50);
    }


    function initializeGoogle() {
        if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR")) {
            setAuthStatus("Add your existing Google Web Client ID to config.js.", true);
            return;
        }

        google.accounts.id.initialize({
            client_id: CONFIG.GOOGLE_CLIENT_ID,
            callback: handleCredentialResponse,
            auto_select: true,
            cancel_on_tap_outside: false
        });

        const saved = readSavedSession();
        if (saved) {
            setUserProfile(saved);
            $("google-signin-button")?.classList.add("hidden");
            $("grant-access").classList.add("hidden");
            const cached = window.FM_AUTH_CACHE?.read?.(saved.email);
            if (cached?.token) {
                state.accessToken = cached.token;
                window.FM_CONNECTION_UI?.show("Restoring your session…", "Using your saved Google Sheets connection.");
                hideLogin();
                loadDashboard().finally(() => window.FM_CONNECTION_UI?.hide());
            } else {
                setAuthStatus("Restoring your Sheets connection…");
                window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Restoring your saved Google session.");
                attemptSilentAccess(saved.email);
            }
        } else {
            window.FM_CONNECTION_UI?.hide();
            google.accounts.id.renderButton($("google-signin-button"), {
                theme: "outline", size: "large", text: "signin_with", shape: "rectangular", width: 280
            });
            google.accounts.id.prompt();
        }
    }


    function handleCredentialResponse(response) {
        try {
            state.idTokenPayload = decodeJwtPayload(response.credential);
            saveSession();
            setUserProfile(state.idTokenPayload);
            $("grant-access").classList.remove("hidden");
            setAuthStatus("Signed in. Connecting to Google Sheets…");
            window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Requesting access for this dashboard.");
            attemptSilentAccess(state.idTokenPayload.email);
        } catch (error) {
            console.error(error);
            setAuthStatus("Google sign-in response could not be read.", true);
        }
    }


    function requestSheetAccess(silent = false) {
        if (!state.idTokenPayload && !readSavedSession()) {
            setAuthStatus("Sign in with Google first.", true);
            return;
        }

        window.FM_CONNECTION_UI?.show(
            silent ? "Connecting to Google Sheets…" : "Authorising Google Sheets & Drive…",
            silent ? "Restoring your saved connection." : "Approve access in the Google prompt, then the dashboard will continue automatically."
        );
        acquireAccessToken(silent ? "none" : "consent")
            .then(async () => {
                hideLogin();
                await loadDashboard();
                window.FM_CONNECTION_UI?.hide();
            })
            .catch(error => {
                window.FM_CONNECTION_UI?.hide();
                $("grant-access").classList.remove("hidden");
                if (silent) {
                    setAuthStatus("Your Google account is signed in. Allow Sheets & Drive access to continue.");
                } else {
                    setAuthStatus(error.message || "Google authorization failed. Tap Allow Sheets & Drive access to try again.", true);
                }
            });
    }


    function attemptSilentAccess(email) {
        window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Checking your current Google authorisation.");
        acquireAccessToken("none", email)
            .then(async () => {
                hideLogin();
                await loadDashboard();
                window.FM_CONNECTION_UI?.hide();
            })
            .catch((error) => {
                console.warn("Silent Google Sheets/Drive authorization failed:", error);
                window.FM_CONNECTION_UI?.hide();
                $("grant-access").classList.remove("hidden");
                setAuthStatus("Your Google session is available. Allow Sheets & Drive access to continue.");
            });
    }


    function acquireAccessToken(prompt = "none", email, options = {}) {
        const forceRefresh = Boolean(options.forceRefresh);
        const expectedEmail = email || state.idTokenPayload?.email || readSavedSession()?.email || "";
        if (!forceRefresh) {
            const cached = window.FM_AUTH_CACHE?.read?.(expectedEmail);
            if (cached?.token) {
                state.accessToken = cached.token;
                return Promise.resolve(cached.token);
            }
        }
        if (tokenRequestPromise) return tokenRequestPromise;
        tokenRequestPromise = new Promise((resolve, reject) => {
            let settled = false;
            const finish = (fn, value) => {
                if (settled) return;
                settled = true;
                tokenRequestPromise = null;
                clearTimeout(timeoutId);
                fn(value);
            };
            const timeoutMs = prompt === "none" ? 5000 : 120000;
            const timeoutId = setTimeout(() => finish(reject, new Error("Google authorisation did not complete.")), timeoutMs);
            const tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: CONFIG.GOOGLE_CLIENT_ID,
                scope: CONFIG.OAUTH_SCOPES,
                callback: response => {
                    if (response.error) { finish(reject, new Error(`Google authorization failed: ${response.error}`)); return; }
                    state.accessToken = response.access_token;
                    window.FM_AUTH_CACHE?.write?.(response.access_token, response.expires_in, expectedEmail);
                    finish(resolve, response.access_token);
                },
                error_callback: error => finish(reject, new Error(
                    error?.type === "popup_failed_to_open"
                        ? "Google sign-in popup was blocked. Allow popups for this site and try again."
                        : "Google authorisation was interrupted. Please try again."
                ))
            });
            tokenClient.requestAccessToken({
                prompt,
                login_hint: expectedEmail || undefined
            });
        });
        return tokenRequestPromise;
    }

    async function ensureWriteAccess() {
        const email = state.idTokenPayload?.email || readSavedSession()?.email || "";
        const cached = window.FM_AUTH_CACHE?.read?.(email);
        if (cached?.token) {
            state.accessToken = cached.token;
            return true;
        }
        try {
            window.FM_CONNECTION_UI?.show("Checking Google access…", "Refreshing access before saving this movement.");
            await acquireAccessToken("none", email, {forceRefresh: true});
            window.FM_CONNECTION_UI?.hide();
            return true;
        } catch (silentError) {
            console.warn("Silent token refresh failed before movement save:", silentError);
            try {
                window.FM_CONNECTION_UI?.show("Google permission needed", "Approve the Google access prompt to finish saving this movement.");
                await acquireAccessToken("consent", email, {forceRefresh: true});
                window.FM_CONNECTION_UI?.hide();
                return true;
            } catch (consentError) {
                window.FM_CONNECTION_UI?.hide();
                throw consentError;
            }
        }
    }

    async function loadDashboard() {

        if (!state.accessToken) {
            return;
        }

        setSyncStatus("Syncing with Assets Inventory Ledger...");

        try {
            const ledgerId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
            const metadata = await sheetsGet(`/${encodeURIComponent(ledgerId)}`);
            let sheetTitles = (metadata.sheets || []).map(sheet => sheet.properties.title);

            const requiredSheets = [
                CONFIG.INVENTORY_SHEET_NAME,
                CONFIG.TRANSACTIONS_SHEET_NAME,
                CONFIG.ROUTINE_SHEET_NAME,
                CONFIG.ALERTS_SHEET_NAME,
                CONFIG.ASSET_AVAILABILITY_SHEET_NAME,
                CONFIG.ASSET_OPERATIONS_SHEET_NAME,
                CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME,
                CONFIG.BROKERS_SHEET_NAME
            ];
            const missingSheets = requiredSheets.filter(name => !sheetTitles.includes(name));

            if (missingSheets.length) {
                await createSheets(missingSheets);
                sheetTitles = sheetTitles.concat(missingSheets);
            }

            const [inventoryRows, transactionRows, clientListRows, routineRows, alertRows, availabilityRows, operationRows, adjustmentRows, clientTypeRows, brokerRows] = await Promise.all([
                getValues(ledgerId, CONFIG.INVENTORY_SHEET_NAME),
                getValues(ledgerId, CONFIG.TRANSACTIONS_SHEET_NAME),
                getValues(ledgerId, CONFIG.CLIENT_LIST_SHEET_NAME),
                getValues(ledgerId, CONFIG.ROUTINE_SHEET_NAME),
                getValues(ledgerId, CONFIG.ALERTS_SHEET_NAME),
                getValues(ledgerId, CONFIG.ASSET_AVAILABILITY_SHEET_NAME),
                getValues(ledgerId, CONFIG.ASSET_OPERATIONS_SHEET_NAME),
                getValues(ledgerId, CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME),
                sheetTitles.includes(CONFIG.CLIENT_TYPES_SHEET_NAME)
                    ? getValues(ledgerId, CONFIG.CLIENT_TYPES_SHEET_NAME).catch(error => {
                        console.warn("Client types could not be loaded:", error);
                        return [];
                    })
                    : Promise.resolve([]),
                getValues(ledgerId, CONFIG.BROKERS_SHEET_NAME).catch(error => {
                    console.warn("Broker reference list unavailable:", error);
                    return [];
                })
            ]);
            if (brokerRows.length <= 1) {
                try {
                    await updateValues(ledgerId, CONFIG.BROKERS_SHEET_NAME, [["Broker"], ...FM_BROKERS.DEFAULT_BROKERS.map(name => [name])]);
                } catch (error) { console.warn("Cannot initialise Brokers sheet:", error); }
            }

            await ensureTransactionColumns(ledgerId, transactionRows);
            await ensureOperationalSheetHeaders(ledgerId, CONFIG.ROUTINE_SHEET_NAME, routineRows, FM_ROUTINE_ALERTS.ROUTINE_HEADERS);
            await ensureOperationalSheetHeaders(ledgerId, CONFIG.ALERTS_SHEET_NAME, alertRows, FM_ROUTINE_ALERTS.ALERT_HEADERS);
            await ensureOperationalSheetHeaders(ledgerId, CONFIG.ASSET_AVAILABILITY_SHEET_NAME, availabilityRows, ["Date", "Asset", "Available Count", "Updated By", "Updated At"]);
            await ensureOperationalSheetHeaders(ledgerId, CONFIG.ASSET_OPERATIONS_SHEET_NAME, operationRows, ["Timestamp", "Client", "Movement", "Asset", "Quantity", "User", "Comment", "Image Link", "Inventory Balance", "Operation ID"]);
            await ensureOperationalSheetHeaders(ledgerId, CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME, adjustmentRows, ["Timestamp", "Client", "Movement", "Asset", "Quantity", "User", "Comment", "Image Link", "Inventory Balance", "Target Balance", "Previous Balance", "Adjustment ID"]);
            state.alertHeaders = alertRows[0]?.slice() || FM_ROUTINE_ALERTS.ALERT_HEADERS.slice();

            state.inventory = parseInventory(inventoryRows);
            state.assetAvailability = parseAvailability(availabilityRows);
            state.assetOperationRows = operationRows;
            state.clientAdjustmentRows = adjustmentRows;
            // Inventory Change Log is retired. Remove the old sheet once and
            // keep Asset Inventory as a plain Asset/Balance table.
            if (window.FM_SHEET_WRITER?.removeLegacyChangeLog) {
                try {
                    await FM_SHEET_WRITER.removeLegacyChangeLog({
                        spreadsheetId: ledgerId,
                        accessToken: state.accessToken
                    });
                } catch (legacyError) {
                    console.warn("Legacy Inventory Change Log could not be removed yet:", legacyError);
                }
            }
            state.transactions = [
                ...parseTransactions(transactionRows),
                ...parseTransactions(operationRows).filter(item => item.client || item.asset),
                ...parseTransactions(adjustmentRows).map((item, index) => ({
                    ...item,
                    // Support older adjustment sheets whose Timestamp heading
                    // was accidentally placed in M1 while the values are in A.
                    timestamp: item.timestamp || String(adjustmentRows[index + 1]?.[0] ?? ""),
                    isLedgerAdjustment: true
                })).filter(item => item.client || item.asset)
            ];
            state.clientTypes = FM_DASHBOARD_FEATURES.parseClientTypes(clientTypeRows);
            state.brokers = FM_BROKERS.parseBrokers(brokerRows);
            state.clientListRows = clientListRows.map(row => row.slice());
            state.clients = parseClients(clientListRows);
            state.routines = FM_ROUTINE_ALERTS.parseRoutineRows(routineRows);
            state.alerts = FM_ROUTINE_ALERTS.parseAlertRows(alertRows);
            state.todayLoads = FM_ROUTINE_ALERTS.routineLoadRows(state.routines, new Date());

            state.workbookId = ledgerId;
            state.workbookName = CONFIG.INVENTORY_LEDGER_NAME || "Assets Inventory Ledger";
            $("workbook-name").textContent = state.workbookName;

            state.dashboardInventory = state.inventory.slice();
            state.dashboardTransactions = state.transactions.slice();
            state.dashboardClients = state.clients.slice();
            state.dashboardSource = { kind: "current", id: ledgerId, name: state.workbookName };

            // Refresh the alert snapshot: stale open alerts are removed;
            // attended/ignored history remains and today's candidates are rebuilt.
            try {
                await rebuildAlertSnapshot();
            } catch (alertError) {
                console.warn("Alerts could not be synchronised:", alertError);
            }

            try {
                await loadHistoricalLedgerFiles();
            } catch (archiveError) {
                console.warn("Historical Dashboard snapshots could not be listed:", archiveError);
                state.historicalLedgerFiles = [];
            }

            populateDashboardDataVersionOptions();
            renderDashboard();
            updateAlertNavBadges();

            setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
        } catch (error) {
            console.error(error);
            setSyncStatus(error.message || "Unable to load spreadsheet.", true);
        }
    }


    function parseHistoricalLedgerName(name) {
        const match = String(name || "").trim().match(/^(\d{1,2})_(\d{4})_(\d{4})$/);
        if (!match) return null;

        const week = Number(match[1]);
        const year = Number(match[2]);
        const hhmm = match[3];
        const hour = Number(hhmm.slice(0, 2));
        const minute = Number(hhmm.slice(2));

        if (!Number.isInteger(week) || week < 1 || week > 53 ||
            !Number.isInteger(year) || !Number.isInteger(hour) || hour > 23 ||
            !Number.isInteger(minute) || minute > 59) {
            return null;
        }

        return {
            week,
            year,
            hhmm,
            timeLabel: `${hhmm.slice(0, 2)}:${hhmm.slice(2)}`,
            label: `Week ${week} · ${year} · ${hhmm.slice(0, 2)}:${hhmm.slice(2)}`
        };
    }


    async function loadHistoricalLedgerFiles() {
        const parent = CONFIG.DRIVE_PHOTOS_PARENT_FOLDER_ID;
        if (!parent) {
            state.historicalLedgerFiles = [];
            return [];
        }

        const files = [];
        let pageToken = "";

        do {
            const params = new URLSearchParams({
                q: `'${parent}' in parents and trashed = false`,
                pageSize: "100",
                orderBy: "name desc",
                fields: "nextPageToken,files(id,name,mimeType,modifiedTime,webViewLink)",
                supportsAllDrives: "true",
                includeItemsFromAllDrives: "true"
            });
            if (pageToken) params.set("pageToken", pageToken);

            const data = await fetchJson(`${DRIVE_API}?${params.toString()}`, {
                headers: authHeaders()
            });

            (data.files || []).forEach(file => {
                const parsed = parseHistoricalLedgerName(file.name);
                if (!parsed) return;

                // Historical dashboard snapshots are expected to be Google Sheets.
                // Do not offer other Drive files that the Sheets API cannot read.
                if (file.mimeType !== "application/vnd.google-apps.spreadsheet") return;

                files.push({ ...file, ...parsed });
            });

            pageToken = data.nextPageToken || "";
        } while (pageToken);

        files.sort((a, b) =>
            (b.year - a.year) ||
            (b.week - a.week) ||
            String(b.hhmm).localeCompare(String(a.hhmm)) ||
            String(b.name).localeCompare(String(a.name))
        );

        state.historicalLedgerFiles = files;
        return files;
    }


    function populateDashboardDataVersionOptions() {
        const select = $("dashboard-data-version");
        if (!select) return;

        const currentValue = select.value || "current";
        const archived = state.historicalLedgerFiles || [];

        select.innerHTML = [
            `<option value="current">Current live ledger</option>`,
            ...archived.map(file =>
                `<option value="${escapeAttr(file.id)}">${escapeHtml(file.label)} · ${escapeHtml(file.name)}</option>`
            )
        ].join("");

        if (currentValue === "current" || archived.some(file => file.id === currentValue)) {
            select.value = currentValue;
        } else {
            select.value = "current";
        }

        updateDashboardDataVersionStatus();
    }


    function updateDashboardDataVersionStatus(message) {
        const status = $("dashboard-data-version-status");
        if (!status) return;

        if (message) {
            status.textContent = message;
            return;
        }

        if (state.dashboardSource?.kind === "historical") {
            status.textContent = `Historical snapshot: ${state.dashboardSource.name}`;
        } else {
            status.textContent = "";
        }
    }


    async function handleDashboardDataVersionChange(event) {
        const value = String(event.target?.value || "current");

        if (value === "current") {
            state.dashboardInventory = state.inventory.slice();
            state.dashboardTransactions = state.transactions.slice();
            state.dashboardClients = state.clients.slice();
            state.dashboardSource = {
                kind: "current",
                id: CONFIG.INVENTORY_LEDGER_SHEET_ID,
                name: CONFIG.INVENTORY_LEDGER_NAME || "Current live ledger"
            };
            updateDashboardDataVersionStatus();
            renderDashboardViews();
            return;
        }

        const file = state.historicalLedgerFiles.find(item => item.id === value);
        if (!file) {
            event.target.value = "current";
            return;
        }

        try {
            event.target.disabled = true;
            updateDashboardDataVersionStatus(`Loading ${file.label}…`);

            const metadata = await sheetsGet(`/${encodeURIComponent(file.id)}`);
            const sheetTitles = (metadata.sheets || []).map(sheet => sheet.properties.title);
            const missing = [CONFIG.INVENTORY_SHEET_NAME, CONFIG.TRANSACTIONS_SHEET_NAME]
                .filter(name => !sheetTitles.includes(name));

            if (missing.length) {
                throw new Error(`The archive ${file.name} is missing: ${missing.join(", ")}.`);
            }

            const [inventoryRows, transactionRows, clientRows] = await Promise.all([
                getValues(file.id, CONFIG.INVENTORY_SHEET_NAME),
                getValues(file.id, CONFIG.TRANSACTIONS_SHEET_NAME),
                sheetTitles.includes(CONFIG.CLIENT_LIST_SHEET_NAME)
                    ? getValues(file.id, CONFIG.CLIENT_LIST_SHEET_NAME)
                    : Promise.resolve([])
            ]);

            const dashboardInventory = parseInventory(inventoryRows);
            const dashboardTransactions = parseTransactions(transactionRows);
            let dashboardClients = parseClients(clientRows);

            if (!dashboardClients.length) {
                dashboardClients = [...new Set(
                    dashboardTransactions
                        .map(item => String(item.client || "").trim())
                        .filter(Boolean)
                )].sort((a, b) => a.localeCompare(b));
            }

            state.dashboardInventory = dashboardInventory;
            state.dashboardTransactions = dashboardTransactions;
            state.dashboardClients = dashboardClients;
            state.dashboardSource = {
                kind: "historical",
                id: file.id,
                name: file.name,
                modifiedTime: file.modifiedTime || ""
            };

            updateDashboardDataVersionStatus();
            renderDashboardViews();
        } catch (error) {
            console.error(error);
            event.target.value = state.dashboardSource?.kind === "historical"
                ? state.dashboardSource.id
                : "current";
            updateDashboardDataVersionStatus(error.message || "Unable to load that archive.");
        } finally {
            event.target.disabled = false;
        }
    }


    // Sheet-creation only ever targets the Assets Inventory Ledger -
    // the routine and alerts sheets are maintained in the Assets Inventory Ledger.
    async function createSheets(
        names
    ) {

        const ledgerId =
            CONFIG.INVENTORY_LEDGER_SHEET_ID;


        const requests =
            names.map(
                title => ({

                    addSheet: {

                        properties: {

                            title

                        }

                    }

                })
            );


        await sheetsPost(

            `/${encodeURIComponent(
                ledgerId
            )}:batchUpdate`,

            {
                requests
            }

        );


        if (
            names.includes(
                CONFIG.INVENTORY_SHEET_NAME
            )
        ) {

            await updateValues(

                ledgerId,

                CONFIG.INVENTORY_SHEET_NAME,

                [
                    [
                        "Asset",
                        "Balance"
                    ]
                ]

            );

        }


        if (names.includes(CONFIG.ROUTINE_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.ROUTINE_SHEET_NAME, [FM_ROUTINE_ALERTS.ROUTINE_HEADERS]);
        }

        if (names.includes(CONFIG.ALERTS_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.ALERTS_SHEET_NAME, [FM_ROUTINE_ALERTS.ALERT_HEADERS]);
        }
        if (names.includes(CONFIG.ASSET_AVAILABILITY_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.ASSET_AVAILABILITY_SHEET_NAME, [["Date", "Asset", "Available Count", "Updated By", "Updated At"]]);
        }
        if (names.includes(CONFIG.ASSET_OPERATIONS_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.ASSET_OPERATIONS_SHEET_NAME, [["Timestamp", "Client", "Movement", "Asset", "Quantity", "User", "Comment", "Image Link", "Inventory Balance", "Operation ID"]]);
        }
        if (names.includes(CONFIG.BROKERS_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.BROKERS_SHEET_NAME, [["Broker"], ...FM_BROKERS.DEFAULT_BROKERS.map(name => [name])]);
        }
        if (names.includes(CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME)) {
            await updateValues(ledgerId, CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME, [["Timestamp", "Client", "Movement", "Asset", "Quantity", "User", "Comment", "Image Link", "Inventory Balance", "Target Balance", "Previous Balance", "Adjustment ID"]]);
        }

        if (
            names.includes(
                CONFIG.TRANSACTIONS_SHEET_NAME
            )
        ) {

            await updateValues(

                ledgerId,

                CONFIG.TRANSACTIONS_SHEET_NAME,

                [
                    [
                        "Timestamp",
                        "Client",
                        "Movement",
                        "Asset",
                        "Quantity",
                        "User",
                        "Comment",
                        "Image Link",
                        "Inventory Balance",
                        "Broker"
                    ]
                ]

            );

        }

    }


    async function ensureOperationalSheetHeaders(ledgerId, sheetName, rows, expectedHeaders) {
        if (!rows.length) {
            await updateValues(ledgerId, sheetName, [expectedHeaders]);
            rows.push(expectedHeaders.slice());
            return;
        }
        const existing = rows[0].slice();
        const normalized = existing.map(normalizeHeader);
        const missing = expectedHeaders.filter(header => !normalized.includes(normalizeHeader(header)));
        if (missing.length) {
            existing.push(...missing);
            await updateValues(ledgerId, sheetName, [existing]);
            rows[0] = existing;
        }
    }

    function alertRowFromEntity(alert, status) {
        const headers = state.alertHeaders.length ? state.alertHeaders : FM_ROUTINE_ALERTS.ALERT_HEADERS;
        const normalized = headers.map(normalizeHeader);
        const row = Array.isArray(alert.rawRow) ? headers.map((_, index) => alert.rawRow[index] ?? "") : new Array(headers.length).fill("");
        const put = (aliases, value) => {
            const wanted = aliases.map(normalizeHeader);
            const index = normalized.findIndex(item => wanted.includes(item));
            if (index >= 0) row[index] = value ?? "";
        };
        put(["alert key","key"], alert.key);
        put(["created at","timestamp"], alert.createdAt || new Date().toISOString());
        put(["alert type","type"], alert.type);
        put(["status"], status || alert.status || "OPEN");
        put(["client"], alert.client); put(["asset"], alert.asset);
        put(["required qty","required"], alert.requiredQty); put(["available qty","available"], alert.availableQty);
        put(["shortfall","short by"], alert.shortfall); put(["destination","location"], alert.destination);
        put(["message","alert"], alert.message); put(["attended at"], alert.attendedAt);
        put(["attended by","attended user"], alert.attendedBy); put(["resolution comment","resolution note"], alert.resolutionComment);
        put(["routine key"], alert.routineKey); put(["ignored at","ignore at"], alert.ignoredAt);
        put(["ignored by","ignored user"], alert.ignoredBy); put(["ignore comment","ignore note"], alert.ignoreComment);
        return row;
    }

    async function rebuildAlertSnapshot() {
        const candidates = FM_ROUTINE_ALERTS.buildAlertCandidates({
            inventory: state.inventory,
            transactions: state.transactions,
            routines: state.routines,
            clients: state.clients,
            date: new Date()
        });
        const retainedMap = new Map();
        state.alerts.filter(alert => alert.status === "ATTENDED" || alert.status === "IGNORED").forEach(alert => {
            if (alert.key) retainedMap.set(alert.key, { ...alert, rawRow: Array.isArray(alert.rawRow) ? alert.rawRow.slice() : undefined });
        });
        const retained = [...retainedMap.values()];
        const retainedKeys = new Set(retained.map(alert => alert.key));
        const createdAt = new Date().toISOString();
        const fresh = candidates.filter(item => item.key && !retainedKeys.has(item.key)).map(item => ({
            rowNumber: 0, key: item.key, createdAt, type: item.type, status: "OPEN", client: item.client || "", asset: item.asset || "",
            requiredQty: item.requiredQty || 0, availableQty: item.availableQty || 0, shortfall: item.shortfall || 0, destination: item.destination || "",
            message: item.message || "", attendedAt: "", attendedBy: "", resolutionComment: "", routineKey: item.routineKey || "", ignoredAt: "", ignoredBy: "", ignoreComment: ""
        }));
        const finalAlerts = [...retained, ...fresh].map((alert, index) => ({ ...alert, rowNumber: index + 2 }));
        const headers = state.alertHeaders.length ? state.alertHeaders : FM_ROUTINE_ALERTS.ALERT_HEADERS;
        const lastCol = columnLetter(headers.length);
        const clearRange = `${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A2:${lastCol}`;
        await sheetsPost(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(clearRange)}:clear`, {});
        const writeRange = `${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A1:${lastCol}${finalAlerts.length + 1}`;
        await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(writeRange)}?valueInputOption=USER_ENTERED`, {
            range: writeRange, majorDimension: "ROWS", values: [headers, ...finalAlerts.map(alert => alertRowFromEntity(alert, alert.status))]
        });
        state.alerts = finalAlerts;
        return finalAlerts;
    }

    // Compatibility aliases for older dashboard code.
    function buildAlertRow(alert, user) {
        return alertRowFromEntity({ ...alert, createdAt: new Date().toISOString() }, "OPEN");
    }
    async function syncGeneratedAlerts() { return rebuildAlertSnapshot(); }
    async function updateAlertStatus(alert, comment) {
        const attendedAt = new Date().toISOString();
        const attendedBy = state.idTokenPayload?.email || readSavedSession()?.email || "Google user";
        Object.assign(alert, { status: "ATTENDED", attendedAt, attendedBy, resolutionComment: comment });
        const headers = state.alertHeaders.length ? state.alertHeaders : FM_ROUTINE_ALERTS.ALERT_HEADERS;
        const lastCol = columnLetter(headers.length);
        const range = `${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A${alert.rowNumber}:${lastCol}${alert.rowNumber}`;
        await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, { range, majorDimension: "ROWS", values: [alertRowFromEntity(alert, "ATTENDED")] });
    }

    async function getValues(
        spreadsheetId,
        sheetName
    ) {

        const range =
            `${quoteSheetName(
                sheetName
            )}!A:AE`;


        const data =
            await sheetsGet(

                `/${encodeURIComponent(
                    spreadsheetId
                )}/values/${encodeURIComponent(
                    range
                )}`

            );


        return data.values || [];

    }


    async function updateValues(
        spreadsheetId,
        sheetName,
        rows
    ) {

        const range =
            `${quoteSheetName(
                sheetName
            )}!A1`;


        return sheetsPut(

            `/${encodeURIComponent(
                spreadsheetId
            )}/values/${encodeURIComponent(
                range
            )}?valueInputOption=USER_ENTERED`,

            {

                range,

                majorDimension:
                    "ROWS",

                values:
                    rows

            }

        );

    }


    function parseInventory(
        rows
    ) {

        if (!rows.length) {
            return [];
        }


        const header =
            rows[0].map(
                normalizeHeader
            );


        const assetIdx =
            findColumn(

                header,

                [
                    "asset",
                    "asset name",
                    "item",
                    "type"
                ]

            );


        const balanceIdx =
            findColumn(

                header,

                [
                    "balance",
                    "current balance",
                    "stock",
                    "quantity"
                ]

            );

        // Optional audit columns. Older workbooks may not have them.
        const verifiedIdx =
            findColumn(
                header,
                ["verified", "verified status", "is verified"]
            );
        const verificationDateIdx =
            findColumn(
                header,
                ["date", "verification date", "verified date", "audit date"]
            );


        if (assetIdx < 0) {
            return [];
        }


        return rows

            .slice(1)

            .map(
                (row, index) => ({

                    rowNumber:
                        index + 2,

                    asset:
                        String(
                            row[assetIdx] ?? ""
                        ).trim(),

                    balance:
                        balanceIdx >= 0
                            ? numericValue(
                                row[balanceIdx]
                            )
                            : 0,

                    verified:
                        verifiedIdx >= 0
                            ? isTruthyVerificationValue(row[verifiedIdx])
                            : false,

                    verificationDate:
                        verificationDateIdx >= 0
                            ? String(row[verificationDateIdx] ?? "").trim()
                            : "",

                    assetColumn:
                        assetIdx + 1,

                    balanceColumn:
                        balanceIdx >= 0
                            ? balanceIdx + 1
                            : 2

                })
            )

            .filter(
                item =>
                    item.asset
            );

    }


    function parseTransactions(rows) {
        if (!rows.length) return [];
        const header = rows[0].map(normalizeHeader);
        const idx = {
            timestamp: findColumn(header, ["timestamp", "date", "datetime"]),
            client: findColumn(header, ["client", "client name"]),
            movement: findColumn(header, ["movement", "type", "direction"]),
            asset: findColumn(header, ["asset", "asset name", "item"]),
            quantity: findColumn(header, ["quantity", "qty"]),
            user: findColumn(header, ["user", "entered by", "email"]),
            comment: findColumn(header, ["comment", "comments", "notes"]),
            image: findColumn(header, ["image link", "image", "photo", "picture", "photo link", "attachment", "drive link"]),
            inventoryBalance: findColumn(header, ["inventory balance", "warehouse balance", "balance after movement"]),
            broker: findColumn(header, ["broker", "broker name"])
        };

        return rows.slice(1).map(row => ({
            timestamp: idx.timestamp >= 0 ? row[idx.timestamp] ?? "" : "",
            client: idx.client >= 0 ? row[idx.client] ?? "" : "",
            movement: idx.movement >= 0 ? row[idx.movement] ?? "" : "",
            asset: idx.asset >= 0 ? row[idx.asset] ?? "" : "",
            quantity: idx.quantity >= 0 ? numericValue(row[idx.quantity]) : 0,
            user: idx.user >= 0 ? row[idx.user] ?? "" : "",
            comment: idx.comment >= 0 ? String(row[idx.comment] ?? "").trim() : "",
            image: idx.image >= 0 ? String(row[idx.image] ?? "").trim() : "",
            inventoryBalance: idx.inventoryBalance >= 0 ? numericValue(row[idx.inventoryBalance]) : null,
            broker: idx.broker >= 0 ? String(row[idx.broker] ?? "").trim() : ""
        })).filter(item => item.asset || item.client);
    }

    function normalizeSheetDateKey(value) {
        const text = String(value ?? "").trim();
        if (!text) return "";
        const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
        if (iso) return iso[1];
        const dmy = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
        if (dmy) return `${dmy[3]}-${String(dmy[2]).padStart(2, "0")}-${String(dmy[1]).padStart(2, "0")}`;
        const serial = Number(text);
        if (Number.isFinite(serial) && serial > 20000 && serial < 80000) {
            const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
            return `${date.getUTCFullYear()}-${String(date.getUTCMonth()+1).padStart(2,"0")}-${String(date.getUTCDate()).padStart(2,"0")}`;
        }
        return "";
    }

    function todayDateKey() {
        const now = new Date();
        return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}`;
    }

    function parseAvailability(rows) {
        if (!Array.isArray(rows) || !rows.length) return [];
        const header = rows[0].map(normalizeHeader);
        const dateIdx = findColumn(header, ["date", "day"]);
        const assetIdx = findColumn(header, ["asset", "asset type", "item"]);
        const countIdx = findColumn(header, ["available count", "available", "available to send", "quantity"]);
        const userIdx = findColumn(header, ["updated by", "user", "entered by"]);
        const updatedIdx = findColumn(header, ["updated at", "timestamp", "last updated"]);
        if (dateIdx < 0 || assetIdx < 0 || countIdx < 0) return [];
        return rows.slice(1).map((row, index) => ({
            rowNumber: index + 2,
            date: normalizeSheetDateKey(row[dateIdx]),
            asset: String(row[assetIdx] ?? "").trim(),
            availableCount: row[countIdx] === "" || row[countIdx] == null ? null : numericValue(row[countIdx]),
            updatedBy: String(row[userIdx] ?? "").trim(),
            updatedAt: String(row[updatedIdx] ?? "").trim()
        })).filter(item => item.date && item.asset && item.availableCount !== null && Number.isFinite(item.availableCount));
    }

    function getTodayAvailability(asset) {
        const key = String(asset || "").trim().toLowerCase();
        const today = todayDateKey();
        const matches = (state.assetAvailability || []).filter(row => row.date === today && row.asset.toLowerCase() === key);
        return matches.length ? matches[matches.length - 1].availableCount : null;
    }

    async function ensureTransactionColumns(ledgerId, rows) {
        const header = rows.length
            ? rows[0].slice()
            : ["Timestamp", "Client", "Movement", "Asset", "Quantity", "User"];

        const normalized = header.map(normalizeHeader);
        let changed = false;

        if (!normalized.some(h => ["comment", "comments", "notes"].includes(h))) {
            header.push("Comment");
            normalized.push("comment");
            changed = true;
        }
        if (!normalized.some(h => ["image link", "image", "photo", "picture", "photo link", "attachment", "drive link"].includes(h))) {
            header.push("Image Link");
            normalized.push("image link");
            changed = true;
        }
        if (!normalized.some(h => ["inventory balance", "warehouse balance", "balance after movement"].includes(h))) {
            header.push("Inventory Balance");
            normalized.push("inventory balance");
            changed = true;
        }
        if (!normalized.includes("broker")) {
            header.push("Broker");
            normalized.push("broker");
            changed = true;
        }

        if (changed) {
            const range = `${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A1`;
            await sheetsPut(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {
                range, majorDimension: "ROWS", values: [header]
            });
            if (rows.length) rows[0] = header;
            else rows.push(header);
        }

        return {header, length: header.length};
    }


    function findMainHeaderRow(
        rows
    ) {

        return rows.findIndex(
            row => {

                const headers =
                    row.map(
                        normalizeHeader
                    );


                const hasClient =
                    headers.includes("cleint") ||
                    headers.includes("client") ||
                    headers.includes("collection client");


                const hasLoadType =
                    headers.includes("load type") ||
                    headers.includes("loadtype");


                const hasArrival =
                    headers.includes("planned arrival") ||
                    headers.includes("plannedarrival");


                return (
                    hasClient &&
                    hasLoadType &&
                    hasArrival
                );

            }
        );

    }


    function parseClients(rows) {
        if (!rows.length) return [];
        const header = rows[0].map(normalizeHeader);
        const clientIdx = findColumn(header, ["client name", "client", "name"]);
        if (clientIdx < 0) return [];
        const clients = new Set();
        rows.slice(1).forEach(row => { const value = String(row[clientIdx] ?? "").trim(); if (value) clients.add(value); });
        return [...clients].sort((a,b)=>a.localeCompare(b));
    }

    function findClientListColumns(rows) {
        const header = Array.isArray(rows?.[0]) ? rows[0].map(normalizeHeader) : [];
        const clientIndex = findColumn(header, ["client name","client","name"]);
        const reportIndex = findColumn(header, ["reports","report","report link","google sheet","google sheets","sheet","sheet link","report spreadsheet","spreadsheet link"]);
        return {header,clientIndex,reportIndex};
    }

    function findClientListRowNumber(rows,client,clientIndex) {
        if(!Array.isArray(rows)||clientIndex<0)return-1;
        const wanted=String(client||"").trim().toLowerCase();
        for(let i=1;i<rows.length;i+=1){if(String(rows[i]?.[clientIndex]??"").trim().toLowerCase()===wanted)return i+1;}
        return -1;
    }

    function spreadsheetIdFromLink(value){
        const text=String(value||"").trim();
        const match=text.match(/spreadsheets\/d\/([a-zA-Z0-9_-]+)/)||text.match(/^([a-zA-Z0-9_-]{20,})$/);
        return match?match[1]:"";
    }

    function clientReportUrl(spreadsheetId){return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/edit`;}

    function selectedClientExportAssets(){
        return [...document.querySelectorAll('#client-ledger-export-assets input[data-client-asset]:checked')].map(input=>String(input.dataset.clientAsset||"").trim()).filter(Boolean);
    }

    function renderClientLedgerExportAssets(client){
        const container=$("client-ledger-export-assets");
        if(!container)return;
        const currentAsset=$("client-details-filter-asset")?.value||"";
        const assets=[...new Set(getDashboardTransactions().filter(item=>String(item.client||"").trim().toLowerCase()===String(client||"").trim().toLowerCase()).map(item=>String(item.asset||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
        if(!assets.length){container.innerHTML='<span class="muted">No asset types are available for this client.</span>';return;}
        const defaultAll=!currentAsset;
        container.innerHTML=assets.map(asset=>`<label class="client-ledger-asset-option"><input type="checkbox" data-client-asset="${escapeAttr(asset)}" ${defaultAll||asset===currentAsset?"checked":""}><span>${escapeHtml(asset)}</span></label>`).join("");
        updateClientLedgerExportAssetSelectionUi();
    }

    function updateClientLedgerExportAssetSelectionUi(){
        const boxes=[...document.querySelectorAll('#client-ledger-export-assets input[data-client-asset]')];
        const selected=boxes.filter(box=>box.checked).length;
        const button=$("client-ledger-export-select-all");
        if(button)button.textContent=boxes.length&&selected===boxes.length?"Clear all":"Select all";
    }
    function handleClientLedgerAssetSelectionChange(){updateClientLedgerExportAssetSelectionUi();}
    function selectAllClientLedgerExportAssets(){
        const boxes=[...document.querySelectorAll('#client-ledger-export-assets input[data-client-asset]')];
        const allSelected=boxes.length>0&&boxes.every(box=>box.checked);
        boxes.forEach(box=>{box.checked=!allSelected;});
        updateClientLedgerExportAssetSelectionUi();
    }


    function isTodayTransaction(
        timestamp
    ) {

        if (!timestamp) {
            return false;
        }


        const parsed =
            new Date(timestamp);


        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            return (
                String(
                    timestamp
                ).slice(0, 10) ===
                new Date()
                    .toISOString()
                    .slice(0, 10)
            );

        }


        const now =
            new Date();


        return (

            parsed.getFullYear() ===
                now.getFullYear() &&

            parsed.getMonth() ===
                now.getMonth() &&

            parsed.getDate() ===
                now.getDate()

        );

    }


    function openClientDetails(client) {
        $("client-details-title").textContent = client;

        const dateSelect = $("client-details-filter-date");
        const assetSelect = $("client-details-filter-asset");
        populateClientDetailsAssetFilter(client);

        const render = () => {
            const dateFilter = dateSelect?.value || "";
            const assetFilter = assetSelect?.value || "";

            const transactions = getDashboardTransactions()
                .filter(item =>
                    String(item.client || "").trim().toLowerCase() === String(client).trim().toLowerCase() &&
                    (!dateFilter || transactionDateKey(item.timestamp) === dateFilter) &&
                    (!assetFilter || String(item.asset || "").trim() === assetFilter)
                );

            const assets = getClientAssetSummary(transactions);

            $("client-details-subtitle").textContent =
                `${assets.length} asset type${assets.length === 1 ? "" : "s"} · ${transactions.length} transaction${transactions.length === 1 ? "" : "s"}`;

            $("client-details-summary").innerHTML = assets.length
                ? assets.map(item => `<button type="button"
                    class="client-asset-summary-row ${
                        item.net < 0 ? "asset-net-negative" : item.net > 0 ? "asset-net-positive" : "asset-net-zero"
                    }" data-client-asset="${escapeAttr(item.asset)}">
                    <span class="client-asset-name">${escapeHtml(item.asset)}</span>
                    <span class="client-asset-quantity">${formatSignedNumber(item.net)}</span>
                    <span class="client-asset-meta">Sent − Received</span>
                </button>`).join("")
                : `<div class="client-assets-empty"><strong>No asset movements match these filters</strong><span>Try another day or asset type.</span></div>`;

            renderClientAssetGraph(client, assets[0]?.asset || "");

            const groups = FM_TRANSACTION_DISPLAY.groupTransactions(transactions);
            state.clientDetailsRows = groups;
            state.clientDetailsVisibleCount = 5;
            renderTransactionSummary($("client-details-transaction-summary"), transactions, "Client totals");
            renderClientDetailsTransactionList();
            $("client-details-subtitle").textContent = `${assets.length} asset type${assets.length === 1 ? "" : "s"} · ${groups.length} grouped movement${groups.length === 1 ? "" : "s"} · ${transactions.length} ledger line${transactions.length === 1 ? "" : "s"}`;
        };

        if (dateSelect) dateSelect.onchange = render;
        if (assetSelect) assetSelect.onchange = render;

        // Start the details modal on the same day filter used by the dashboard.
        if (dateSelect && $("dashboard-filter-date")) dateSelect.value = $("dashboard-filter-date").value || "";
        render();
        $("client-details-modal").classList.remove("hidden");
    }

    function renderClientDetailsTransactionList() {
        const container = $("client-details-transactions");
        if (!container) return;
        const groups = state.clientDetailsRows || [];
        if (!groups.length) {
            container.innerHTML = `<div class="client-assets-empty"><strong>No asset transactions recorded for this client.</strong><span>Try another day or asset type.</span></div>`;
            return;
        }
        const visibleCount = Math.min(groups.length, Math.max(5, state.clientDetailsVisibleCount || 5));
        const cards = groups.slice(0, visibleCount)
            .map((group, index) => renderTransactionCard(group, index, "data-client-info-index")).join("");
        const remaining = groups.length - visibleCount;
        const pager = remaining > 0
            ? `<div class="client-details-load-more-wrap"><span>Showing ${visibleCount.toLocaleString()} of ${groups.length.toLocaleString()} movements</span><button type="button" class="secondary" data-client-load-more>Load 5 more</button></div>`
            : groups.length > 5
                ? `<div class="client-details-load-more-wrap"><span>Showing all ${groups.length.toLocaleString()} movements</span></div>`
                : "";
        container.innerHTML = cards + pager;
    }

    function populateClientDetailsAssetFilter(client) {
        const select = $("client-details-filter-asset");
        if (!select) return;

        const current = select.value || "";
        const assets = [...new Set(getDashboardTransactions()
            .filter(t => String(t.client || "").trim().toLowerCase() === String(client).trim().toLowerCase())
            .map(t => String(t.asset || "").trim())
            .filter(Boolean))]
            .sort((a,b) => a.localeCompare(b));

        select.innerHTML = `<option value="">All asset types</option>${assets.map(a => `<option value="${escapeAttr(a)}">${escapeHtml(a)}</option>`).join("")}`;
        if (assets.includes(current)) select.value = current;
    }

    function getClientAssetSummary(transactions) {

        const map = new Map();

        transactions.forEach(item => {

            const asset =
                String(item.asset || "").trim();

            if (!asset) {
                return;
            }

            const key =
                asset.toLowerCase();

            if (!map.has(key)) {
                map.set(
                    key,
                    {
                        asset,
                        sent: 0,
                        received: 0,
                        normalisation: 0
                    }
                );
            }

            const entry =
                map.get(key);

            const movement =
                String(item.movement || "")
                    .trim()
                    .toUpperCase();

            const quantity =
                Number(item.quantity) || 0;

            if (movement === "SENT") {
                entry.sent += quantity;
            }

            if (movement === "RECEIVED") {
                entry.received += quantity;
            }
            if (movement === "NORMALISATION") {
                entry.normalisation += quantity;
            }
        });

        return [...map.values()]
            .map(
                item => ({
                    ...item,
                    net:
                        item.sent - item.received + (item.normalisation || 0)
                })
            )
            .sort(
                (a, b) =>
                    a.asset.localeCompare(b.asset)
            );
    }


    function syncGraphRangeControls(group) {
        document.querySelectorAll(`[data-graph-range-group="${group}"] [data-graph-range]`).forEach(button => {
            const active = button.dataset.graphRange === graphRanges[group];
            button.classList.toggle("is-active", active);
            button.setAttribute("aria-pressed", String(active));
        });
    }

    function graphLocalDate(key) {
        // Noon avoids local timezone / daylight-saving gaps around midnight.
        return new Date(`${key}T12:00:00`);
    }

    function graphDateKey(date) {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
    }

    function graphAddDays(key, days) {
        const date = graphLocalDate(key);
        date.setDate(date.getDate() + days);
        return graphDateKey(date);
    }

    function graphBucketKey(key, granularity) {
        if (granularity === "month") return key.slice(0, 7) + "-01";
        if (granularity === "week") {
            const date = graphLocalDate(key);
            date.setDate(date.getDate() - ((date.getDay() + 6) % 7));
            return graphDateKey(date);
        }
        return key;
    }

    function graphAdvanceBucket(key, granularity) {
        if (granularity === "day") return graphAddDays(key, 1);
        if (granularity === "week") return graphAddDays(key, 7);
        const date = graphLocalDate(key);
        date.setMonth(date.getMonth() + 1, 1);
        return graphDateKey(date);
    }

    function graphPeriodLabel(key, granularity) {
        if (granularity === "month") return graphLocalDate(key).toLocaleDateString("en-GB", {month: "short", year: "2-digit"});
        return formatGraphDate(key);
    }

    function graphWindow(daily, group) {
        const keys = daily.map(([date]) => date).filter(Boolean).sort();
        const today = graphDateKey(new Date());
        const sourceIsHistorical = state.dashboardSource?.kind === "historical";
        const end = sourceIsHistorical && keys.length ? keys[keys.length - 1] : today;
        const range = graphRanges[group];
        const days = range === "7d" ? 7 : range === "30d" ? 30 : range === "90d" ? 90 : null;
        const start = days ? graphAddDays(end, 1 - days) : (keys[0] || end);
        const granularity = range === "90d" ? "week"
            : range === "all" && graphLocalDate(end) - graphLocalDate(start) > 180 * 86400000 ? "month"
            : range === "all" && graphLocalDate(end) - graphLocalDate(start) > 45 * 86400000 ? "week"
            : "day";
        return {start, end, granularity, range};
    }

    function graphWindowSummary(container, windowInfo, points, kind) {
        const info = container?.closest(".details-section")?.querySelector("[data-graph-period]");
        if (!info) return;
        const label = windowInfo.granularity === "week" ? "Weekly" : windowInfo.granularity === "month" ? "Monthly" : "Daily";
        const startYear = graphLocalDate(windowInfo.start).getFullYear();
        const endYear = graphLocalDate(windowInfo.end).getFullYear();
        const withYear = windowInfo.range === "all" || startYear !== endYear;
        const dateLabel = key => withYear ? graphLocalDate(key).toLocaleDateString("en-GB", {day: "numeric", month: "short", year: "numeric"}) : formatGraphDate(key);
        info.textContent = `${dateLabel(windowInfo.start)} – ${dateLabel(windowInfo.end)} · ${label} ${kind}`;
        const scrollHint = container?.closest(".details-section")?.querySelector("[data-graph-hint]");
        if (scrollHint) scrollHint.textContent = "Swipe to explore · Tap a point for details";
    }

    // For movement graphs: sum movements in calendar buckets. Fill missing days
    // with 0 so a week's view is always an actual seven-day window.
    function graphMovementPoints(daily, group) {
        const windowInfo = graphWindow(daily, group);
        const map = new Map();
        daily.forEach(([date, value]) => {
            if (date < windowInfo.start || date > windowInfo.end) return;
            const key = graphBucketKey(date, windowInfo.granularity);
            const current = map.get(key) || {sent: 0, received: 0, discarded: 0};
            current.sent += value.sent || 0;
            current.received += value.received || 0;
            current.discarded += value.discarded || 0;
            map.set(key, current);
        });
        const points = [];
        let key = graphBucketKey(windowInfo.start, windowInfo.granularity);
        const last = graphBucketKey(windowInfo.end, windowInfo.granularity);
        for (let guard = 0; key <= last && guard < 1800; guard++) {
            points.push([key, map.get(key) || {sent: 0, received: 0, discarded: 0}]);
            key = graphAdvanceBucket(key, windowInfo.granularity);
        }
        return {points, windowInfo, hasMovements: [...map.values()].some(v => v.sent || v.received || v.discarded)};
    }

    // Stock graphs use the last known closing quantity per bucket, NOT a sum
    // of balances. Carry forward on dates without any transactions.
    function graphQuantityPoints(daily, group) {
        const windowInfo = graphWindow(daily, group);
        let carry = 0;
        const map = new Map();
        daily.forEach(([date, value]) => {
            if (date < windowInfo.start) carry = value;
            else if (date <= windowInfo.end) map.set(graphBucketKey(date, windowInfo.granularity), value);
        });
        const points = [];
        let key = graphBucketKey(windowInfo.start, windowInfo.granularity);
        const last = graphBucketKey(windowInfo.end, windowInfo.granularity);
        for (let guard = 0; key <= last && guard < 1800; guard++) {
            if (map.has(key)) carry = map.get(key);
            points.push([key, carry]);
            key = graphAdvanceBucket(key, windowInfo.granularity);
        }
        return {points, windowInfo};
    }

    function graphChartWidth(pointCount, windowInfo) {
        const pointWidth = windowInfo.granularity === "day" ? 70 : 86;
        return Math.max(540, Math.min(30000, pointCount * pointWidth));
    }

    function graphXAxisLabels(points, xFor, y, windowInfo) {
        // Skip overlapping labels on long timelines. All data points remain accessible.
        const every = points.length > 20 ? Math.ceil(points.length / 14) : points.length > 12 ? 2 : 1;
        return points.map(([date], i) => i % every === 0 || i === points.length - 1
            ? `<text x="${xFor(i).toFixed(1)}" y="${y}" text-anchor="middle" class="graph-label">${escapeHtml(graphPeriodLabel(date, windowInfo.granularity))}</text>`
            : "").join("");
    }

    function niceGraphMax(value, segments = 4) {
        const maximum = Math.max(Number(value) || 0, 1);
        const roughStep = maximum / segments;
        const magnitude = Math.pow(10, Math.floor(Math.log10(roughStep)));
        const normalized = roughStep / magnitude;
        const niceNormalized = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
        return niceNormalized * magnitude * segments;
    }

    function bindGraphTooltips(scope) {
        if (!scope || scope.dataset.graphTooltipBound === "1") return;
        const scroll = scope.closest ? scope.closest(".client-graph-scroll") : null;
        if (!scroll) return;
        scroll.dataset.graphTooltipBound = "1";
        let tooltip = scroll.querySelector(".graph-hover-tooltip");
        if (!tooltip) {
            tooltip = document.createElement("div");
            tooltip.className = "graph-hover-tooltip";
            tooltip.setAttribute("role", "status");
            tooltip.setAttribute("aria-live", "polite");
            scroll.appendChild(tooltip);
        }
        let hideTimer = null;
        const hide = () => {
            window.clearTimeout(hideTimer);
            tooltip.classList.remove("is-visible");
            tooltip.textContent = "";
        };
        const show = (target, touch = false) => {
            const message = target?.dataset?.graphTooltip || "";
            if (!message) return;
            tooltip.textContent = message;
            tooltip.classList.add("is-visible");
            const sr = scroll.getBoundingClientRect();
            const pr = target.getBoundingClientRect();
            const x = pr.left - sr.left + pr.width / 2;
            const y = pr.top - sr.top;
            tooltip.style.left = `${scroll.scrollLeft + Math.max(16, Math.min(scroll.clientWidth - 16, x))}px`;
            tooltip.style.top = `${Math.max(28, y - 8)}px`;
            window.clearTimeout(hideTimer);
            hideTimer = window.setTimeout(hide, touch ? 3200 : 2400);
        };
        scope.addEventListener("pointerover", event => {
            const target = event.target.closest?.("[data-graph-tooltip]");
            if (target && event.pointerType !== "touch") show(target, false);
        });
        scope.addEventListener("pointerout", event => {
            if (event.pointerType === "touch") return;
            const target = event.target.closest?.("[data-graph-tooltip]");
            if (target && !target.contains(event.relatedTarget)) hide();
        });
        scope.addEventListener("pointerdown", event => {
            const target = event.target.closest?.("[data-graph-tooltip]");
            if (!target) return;
            show(target, event.pointerType === "touch");
        }, {passive:true});
        scroll.addEventListener("scroll", hide, {passive:true});
        window.addEventListener("resize", hide, {passive:true});
    }

    function graphPointHit(x, y, message, className) {
        return `<circle cx="${Number(x).toFixed(1)}" cy="${Number(y).toFixed(1)}" r="13" class="graph-point-hit ${className || ""}" data-graph-tooltip="${escapeAttr(message)}"></circle>`;
    }

    function renderClientAssetGraph(
        client,
        asset
    ) {
        currentClientGraph = {client, asset};
        syncGraphRangeControls("client");

        const title =
            $("client-graph-title");

        const subtitle =
            $("client-graph-subtitle");

        const graph =
            $("client-sent-graph");

        if (!asset) {

            title.textContent =
                "Sent vs received";

            subtitle.textContent =
                "Select an asset type to see its sent and received history.";

            graph.innerHTML =
                `<div class="client-graph-empty">
                    No asset transaction history to graph.
                 </div>`;

            return;
        }

        title.textContent =
            `${asset} · sent vs received`;

        subtitle.textContent =
            "Sent and received quantities from the Asset Transactions ledger.";

        const rows =
            getDashboardTransactions().filter(
                item =>
                    String(item.client).trim().toLowerCase() ===
                    String(client).trim().toLowerCase() &&
                    String(item.asset).trim().toLowerCase() ===
                    String(asset).trim().toLowerCase() &&
                    ["SENT", "RECEIVED"].includes(
                        String(item.movement).trim().toUpperCase()
                    )
            );

        const daily = new Map();

        rows.forEach(item => {

            const date =
                transactionDateKey(
                    item.timestamp
                );

            if (!date) {
                return;
            }

            if (!daily.has(date)) {
                daily.set(date, {
                    sent: 0,
                    received: 0
                });
            }

            const entry = daily.get(date);
            const quantity = Number(item.quantity) || 0;
            const movement = String(item.movement).trim().toUpperCase();

            if (movement === "SENT") {
                entry.sent += quantity;
            } else if (movement === "RECEIVED") {
                entry.received += quantity;
            }
        });

        const history = [...daily.entries()]
                .sort((a, b) => a[0].localeCompare(b[0]));

        if (!history.length) {

            graph.innerHTML =
                `<div class="client-graph-empty">
                    No sent or received transactions recorded for ${escapeHtml(asset)}.
                 </div>`;

            return;
        }

        const {points, windowInfo, hasMovements} = graphMovementPoints(history, "client");
        graphWindowSummary(graph, windowInfo, points, "movements");
        if (!hasMovements) {
            graph.innerHTML = `<div class="client-graph-empty"><strong>No movements in this period</strong><span>Choose a longer time range to see older activity.</span></div>`;
            return;
        }
        const width = graphChartWidth(points.length, windowInfo);

        const height = 300;
        const left = 56;
        const right = 24;
        const top = 42;
        const bottom = 58;
        const chartWidth = width - left - right;
        const chartHeight = height - top - bottom;

        const rawMaxValue =
            Math.max(
                ...points.flatMap(([, value]) => [value.sent, value.received]),
                1
            );
        const maxValue = niceGraphMax(rawMaxValue, 4);

        const yTicks = 4;
        const yGrid = Array.from(
            { length: yTicks + 1 },
            (_, index) => {
                const value =
                    maxValue * (1 - index / yTicks);

                const y =
                    top + (chartHeight * index / yTicks);

                return { value, y };
            }
        );

        const xFor = index =>
            points.length === 1
                ? left + chartWidth / 2
                : left + (index / (points.length - 1)) * chartWidth;

        const yFor = value =>
            top + chartHeight -
            (value / maxValue) * chartHeight;

        const makePath = key =>
            points.map(
                ([, value], index) =>
                    `${index === 0 ? "M" : "L"} ${xFor(index).toFixed(1)} ${yFor(value[key]).toFixed(1)}`
            ).join(" ");

        const sentPath = makePath("sent");
        const receivedPath = makePath("received");

        const gridLines = yGrid.map(tick => `
            <line
                x1="${left}"
                y1="${tick.y.toFixed(1)}"
                x2="${width - right}"
                y2="${tick.y.toFixed(1)}"
                class="graph-grid"
            ></line>
            <text
                x="${left - 9}"
                y="${(tick.y + 4).toFixed(1)}"
                text-anchor="end"
                class="graph-y-label"
            >
                ${formatNumber(Math.round(tick.value))}
            </text>
        `).join("");

        const xLabels = graphXAxisLabels(points, xFor, height - 20, windowInfo);

        const pointsMarkup = points.map(([date, value], index) => {
            const x = xFor(index);
            const sentY = yFor(value.sent);
            const receivedY = yFor(value.received);
            const label = graphPeriodLabel(date, windowInfo.granularity);
            return `
                <g>
                    <circle cx="${x.toFixed(1)}" cy="${sentY.toFixed(1)}" r="4.5" class="graph-point-sent"></circle>
                    ${graphPointHit(x, sentY, `${label} · Sent: ${formatNumber(value.sent)}`, "graph-point-hit-sent")}
                    <circle cx="${x.toFixed(1)}" cy="${receivedY.toFixed(1)}" r="4.5" class="graph-point-received"></circle>
                    ${graphPointHit(x, receivedY, `${label} · Received: ${formatNumber(value.received)}`, "graph-point-hit-received")}
                </g>`;
        }).join("");

        graph.innerHTML = `
            <div class="client-graph-legend">
                <span class="graph-legend-item">
                    <span class="graph-legend-dot graph-legend-sent"></span>
                    Sent
                </span>
                <span class="graph-legend-item">
                    <span class="graph-legend-dot graph-legend-received"></span>
                    Received
                </span>
            </div>

            <div class="client-graph-scroll">
                <svg
                    class="client-graph-svg"
                    style="--graph-min-width:${width}px"
                    viewBox="0 0 ${width} ${height}"
                    role="img"
                    aria-label="Sent versus received history for ${escapeHtml(asset)}"
                >
                    ${gridLines}

                    <line
                        x1="${left}"
                        y1="${top + chartHeight}"
                        x2="${width - right}"
                        y2="${top + chartHeight}"
                        class="graph-axis"
                    ></line>

                    <path
                        d="${sentPath}"
                        class="graph-line graph-line-sent"
                        fill="none"
                    ></path>

                    <path
                        d="${receivedPath}"
                        class="graph-line graph-line-received"
                        fill="none"
                    ></path>

                    ${pointsMarkup}
                    ${xLabels}
                </svg>
            </div>
        `;
        bindGraphTooltips(graph.querySelector(".client-graph-svg"));
    }

    function transactionDateKey(timestamp) {

        if (!timestamp) {
            return "";
        }

        const parsed =
            new Date(timestamp);

        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            const match =
                String(timestamp)
                    .match(/(\d{4}-\d{2}-\d{2})/);

            return match
                ? match[1]
                : "";
        }

        const year =
            parsed.getFullYear();

        const month =
            String(
                parsed.getMonth() + 1
            ).padStart(2, "0");

        const day =
            String(
                parsed.getDate()
            ).padStart(2, "0");

        return `${year}-${month}-${day}`;
    }


    function formatGraphDate(dateKey) {

        const parsed =
            new Date(
                `${dateKey}T00:00:00`
            );

        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {
            return dateKey;
        }

        return parsed.toLocaleDateString(
            "en-GB",
            {
                day: "2-digit",
                month: "short"
            }
        );
    }


    function formatSignedNumber(value) {

        const number =
            Number(value) || 0;

        if (number > 0) {
            return `+${formatNumber(number)}`;
        }

        return formatNumber(number);
    }


    function closeClientDetails() {

        $("client-details-modal")
            .classList
            .add("hidden");

    }


    function formatTransactionTime(
        timestamp
    ) {

        const parsed =
            new Date(timestamp);


        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            return escapeHtml(
                String(
                    timestamp || ""
                )
            );

        }


        return escapeHtml(

            parsed.toLocaleTimeString(
                "en-GB",
                {

                    hour:
                        "2-digit",

                    minute:
                        "2-digit"

                }
            )

        );

    }


    function movementClass(
        movement
    ) {

        const value =
            String(
                movement
            )
                .trim()
                .toUpperCase();


        if (
            value ===
            "RECEIVED"
        ) {

            return "movement-received";

        }


        if (
            value ===
            "SENT"
        ) {

            return "movement-sent";

        }


        if (
            value ===
            "DISCARD"
        ) {

            return "movement-discard";

        }


        return "movement-other";

    }


    function getDashboardFilters() {
        return {
            client: String($("dashboard-filter-client")?.value || "").trim(),
            date: String($("dashboard-filter-date")?.value || "").trim(),
            asset: String($("dashboard-filter-asset")?.value || "").trim()
        };
    }

    function getDashboardInventory() {
        return Array.isArray(state.dashboardInventory)
            ? state.dashboardInventory
            : state.inventory;
    }

    function getDashboardTransactions() {
        return Array.isArray(state.dashboardTransactions)
            ? state.dashboardTransactions
            : state.transactions;
    }

    function getDashboardClients() {
        const source = Array.isArray(state.dashboardClients)
            ? state.dashboardClients
            : state.clients;
        return source.filter(client => client !== "HSC London(Self)");
    }

    function renderDashboardViews() {
        populateDashboardFilters();
        renderWarehouseAssets();
        renderClientCards();
        renderBrokerView();
        renderAudit();
        renderGlobalSearch();
        renderFilteredLoadTables();
        updateDashboardDataVersionStatus();
    }

    function renderBrokerView() {
        window.FM_BROKER_VIEW?.render?.({
            rows: getDashboardTransactions().filter(row => !row.isLedgerAdjustment && !FM_DASHBOARD_FEATURES.isAdjustment(row)),
            brokers: state.brokers
        });
    }

    function filteredDashboardTransactions() {
        const filters = getDashboardFilters();
        return getDashboardTransactions().filter(item =>
            (!filters.client || String(item.client || "").trim() === filters.client) &&
            (!filters.asset || String(item.asset || "").trim() === filters.asset) &&
            (!filters.date || transactionDateKey(item.timestamp) === filters.date)
        );
    }

    function transactionSummaryGroups(transactions) {
        const groups = new Map();
        transactions.forEach(item => {
            const client = String(item.client || "Unknown client").trim() || "Unknown client";
            const asset = String(item.asset || "Unknown asset").trim() || "Unknown asset";
            const movement = String(item.movement || "").trim().toUpperCase();
            const direction = movementLabel(movement) || movement || "Other";
            const key = `${client.toLowerCase()}\u0000${asset.toLowerCase()}\u0000${movement}`;
            if (!groups.has(key)) groups.set(key, {client, asset, movement, direction, quantity: 0});
            groups.get(key).quantity += Number(item.quantity) || 0;
        });
        return [...groups.values()].sort((a, b) =>
            a.client.localeCompare(b.client) ||
            a.asset.localeCompare(b.asset) ||
            a.direction.localeCompare(b.direction)
        );
    }

    function renderTransactionSummary(container, transactions, label = "Matching totals") {
        if (!container) return;
        const groups = transactionSummaryGroups(transactions);
        if (!groups.length) {
            container.innerHTML = "";
            container.classList.add("hidden");
            return;
        }
        const rows = groups.map(group => `<div class="transaction-total-row">
            <strong class="transaction-total-client">${escapeHtml(group.client)}</strong>
            <span class="transaction-total-asset">${escapeHtml(group.asset)}</span>
            <span class="movement-badge ${movementClass(group.movement)}">${escapeHtml(group.direction)}</span>
            <strong class="transaction-total-quantity">${formatNumber(group.quantity)}</strong>
        </div>`).join("");
        container.innerHTML = `<div class="transaction-summary-heading">
            <div class="transaction-summary-heading-copy">
                <div class="eyebrow">${escapeHtml(label)}</div>
                <strong>${formatNumber(transactions.length)} matching transaction${transactions.length === 1 ? "" : "s"}</strong>
            </div>
            <div class="transaction-summary-heading-actions">
                ${label === "Client totals" ? '<button type="button" class="primary transaction-summary-export" data-export-client-ledger>Export ledger</button>' : ""}
                <span class="muted transaction-summary-hint">Grouped by client → asset type → direction</span>
                <button type="button" class="secondary transaction-summary-toggle" data-summary-toggle aria-expanded="false">Expand totals</button>
            </div>
        </div>
        <div class="transaction-summary-body hidden" data-summary-body>
            <div class="transaction-summary-table">
                <div class="transaction-summary-head"><span>Client</span><span>Asset type</span><span>Direction</span><span class="num">Total</span></div>
                ${rows}
            </div>
        </div>`;
        container.classList.remove("hidden");
    }

    const CLIENT_LEDGER_EXCELJS_SRC = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
    let excelJsLoadPromise = null;

    function ensureExcelJS() {
        if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
        if (excelJsLoadPromise) return excelJsLoadPromise;
        excelJsLoadPromise = new Promise((resolve, reject) => {
            const existing = document.querySelector('script[data-client-ledger-exceljs]');
            if (existing) {
                existing.addEventListener("load", () => window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error("ExcelJS loaded without exposing ExcelJS.")), {once:true});
                existing.addEventListener("error", () => reject(new Error("Spreadsheet library failed to load.")), {once:true});
                return;
            }
            const script = document.createElement("script");
            script.src = CLIENT_LEDGER_EXCELJS_SRC;
            script.async = true;
            script.dataset.clientLedgerExceljs = "1";
            script.onload = () => window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error("ExcelJS loaded without exposing ExcelJS."));
            script.onerror = () => reject(new Error("Spreadsheet library failed to load."));
            document.head.appendChild(script);
        });
        return excelJsLoadPromise;
    }

    function openClientLedgerExport() {
        const modal = $("client-ledger-export-modal");
        if (!modal) return;
        const client = $("client-details-title")?.textContent?.trim() || "Unknown client";
        $("client-ledger-export-client").textContent = client;
        const dateFilter = $("client-details-filter-date")?.value || "";
        const dates = getDashboardTransactions()
            .filter(item => String(item.client || "").trim().toLowerCase() === client.toLowerCase())
            .map(item => transactionDateKey(item.timestamp)).filter(Boolean).sort();
        if ($("client-ledger-export-from")) $("client-ledger-export-from").value = dateFilter || dates[0] || "";
        if ($("client-ledger-export-to")) $("client-ledger-export-to").value = dateFilter || dates[dates.length - 1] || "";
        if ($("client-ledger-export-week")) $("client-ledger-export-week").value = isoWeekInputForDateKey(dateFilter || dates[dates.length - 1] || transactionDateKey(new Date()));
        renderClientLedgerExportAssets(client);
        const status = $("client-ledger-export-status");
        if (status) { status.textContent = ""; status.classList.remove("error"); }
        const run = $("run-client-ledger-export");
        if (run) { run.disabled = false; run.textContent = "Save XLSX + Client Sheet"; }
        modal.classList.remove("hidden");
    }

    function closeClientLedgerExport() { $("client-ledger-export-modal")?.classList.add("hidden"); }

    function clientLedgerDirection(movement, quantity) {
        const m = String(movement || "").trim().toUpperCase();
        if (["SENT","OUTBOUND","OUT","SEND"].includes(m)) return "Sent";
        if (["RECEIVED","INBOUND","IN","RECEIVE"].includes(m)) return "Received";
        if (["DISCARDED","DISCARD"].includes(m)) return "Discarded";
        // In client-specific reports, represent a signed ledger correction on
        // the sent/received side that produces the corrected running balance.
        if (["NORMALISATION", "NORMALIZATION", "LEDGER ADJUSTMENT"].includes(m)) {
            if (quantity === undefined || quantity === null || quantity === "") return "Adjustment";
            return Number(quantity) >= 0 ? "Sent" : "Received";
        }
        return movementLabel(movement) || String(movement || "Other");
    }

    function clientLedgerQuantity(item) {
        const quantity = Number(item?.quantity) || 0;
        const movement = String(item?.movement || "").trim().toUpperCase();
        return ["NORMALISATION", "NORMALIZATION", "LEDGER ADJUSTMENT"].includes(movement)
            ? Math.abs(quantity)
            : quantity;
    }

    function exportDateTime(item) {
        const d = new Date(item.timestamp);
        if (Number.isNaN(d.getTime())) return {date:"", time:""};
        const pad = value => String(value).padStart(2, "0");
        return {date:`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`, time:`${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`};
    }

    function makeClientLedgerData(client,allRows,fromKey,toKey,assetFilters=[]){
        const wantedAssets=Array.isArray(assetFilters)?assetFilters.map(v=>String(v).trim()).filter(Boolean):[];
        const wantedSet=new Set(wantedAssets.map(v=>v.toLowerCase()));
        const clientRows=allRows.filter(item=>String(item.client||"").trim().toLowerCase()===String(client||"").trim().toLowerCase());
        const wrap=(item,originalIndex)=>({item,originalIndex,...exportDateTime(item)});
        const allSorted=clientRows.map(wrap).filter(r=>r.date).sort((a,b)=>(new Date(a.item.timestamp).getTime()-new Date(b.item.timestamp).getTime())||(a.originalIndex-b.originalIndex));
        const inPeriod=allSorted.filter(r=>(!fromKey||r.date>=fromKey)&&(!toKey||r.date<=toKey)&&(!wantedSet.size||wantedSet.has(String(r.item.asset||"").trim().toLowerCase())));
        const cumulative=new Map(), cumulativeByRow=new Map();
        for(const row of allSorted){
            const asset=String(row.item.asset||"Unknown asset").trim()||"Unknown asset";
            const qty=clientLedgerQuantity(row.item);
            const direction=clientLedgerDirection(row.item.movement, row.item.quantity);
            const next=(cumulative.get(asset)||0)+(direction==="Sent"?qty:direction==="Received"?-qty:0);
            cumulative.set(asset,next); cumulativeByRow.set(row.item,next); row.asset=asset; row.allTimeDifference=next;
        }
        inPeriod.forEach(row=>{row.asset=String(row.item.asset||"Unknown asset").trim()||"Unknown asset";row.allTimeDifference=cumulativeByRow.get(row.item)||0;});
        const daily=new Map();
        for(const row of inPeriod){
            const key=`${row.date}\u0000${row.asset}`;
            if(!daily.has(key))daily.set(key,{date:row.date,asset:row.asset,sent:0,received:0,difference:0,transactions:0,photoCount:0,allTimeDifference:row.allTimeDifference});
            const d=daily.get(key),qty=clientLedgerQuantity(row.item),direction=clientLedgerDirection(row.item.movement, row.item.quantity);
            if(direction==="Sent"){d.sent+=qty;d.difference+=qty;}else if(direction==="Received"){d.received+=qty;d.difference-=qty;}
            d.transactions+=1;if(row.item.image)d.photoCount+=1;d.allTimeDifference=row.allTimeDifference;
        }
        const dailyRows=[...daily.values()].sort((a,b)=>b.date.localeCompare(a.date)||a.asset.localeCompare(b.asset));
        const transactions=[...inPeriod].sort((a,b)=>(new Date(b.item.timestamp).getTime()-new Date(a.item.timestamp).getTime())||(b.originalIndex-a.originalIndex));
        const periodTotals=inPeriod.reduce((acc,row)=>{const qty=clientLedgerQuantity(row.item),d=clientLedgerDirection(row.item.movement,row.item.quantity);if(d==="Sent")acc.sent+=qty;else if(d==="Received")acc.received+=qty;else if(d==="Discarded")acc.discarded+=qty;acc.transactions+=1;if(row.item.image)acc.photos+=1;return acc;},{sent:0,received:0,discarded:0,transactions:0,photos:0});
        periodTotals.difference=periodTotals.sent-periodTotals.received;
        const assets=[...cumulative.entries()].filter(([asset])=>!wantedSet.size||wantedSet.has(String(asset).toLowerCase())).sort((a,b)=>a[0].localeCompare(b[0])).map(([asset,balance])=>({asset,balance}));
        return{transactions,dailyRows,periodTotals,assets};
    }

    function exportDateGroups(data){
        const groups=new Map();
        for(const row of data.transactions){if(!groups.has(row.date))groups.set(row.date,[]);groups.get(row.date).push(row);}
        return[...groups.entries()].sort((a,b)=>b[0].localeCompare(a[0])).map(([date,rows])=>({date,rows,dailyAssets:data.dailyRows.filter(item=>item.date===date).sort((a,b)=>a.asset.localeCompare(b.asset))}));
    }

    // Adjacent asset rows sharing the same date/time belong to one physical movement.
    // These groups drive PDF rowspans, Excel/Sheets merged cells, and one photo per movement.
    function ledgerTimestampGroups(rows) {
        const groups = [];
        (rows || []).forEach((row, index) => {
            // Ledger exports may merge asset rows only when date/time AND registration match.
            // A blank registration is not reliable evidence of a shared physical movement.
            const registration = FM_TRANSACTION_DISPLAY?.extractVehicleRegistration?.(row?.item?.comment) || "";
            const registrationKey = registration
                ? registration.toUpperCase().replace(/[^A-Z0-9]/g, "")
                : `unregistered-row:${index}`;
            const key = `${row.date || ""}\u0000${row.time || ""}\u0000${registrationKey}`;
            let group = groups[groups.length - 1];
            if (!group || group.key !== key) {
                group = {key, startIndex:index, endIndex:index + 1, count:1, rows:[row], comment:"", images:[], image:""};
                groups.push(group);
            } else {
                group.endIndex = index + 1;
                group.count += 1;
                group.rows.push(row);
            }
        });
        for (const group of groups) {
            group.comment = [...new Set(group.rows.map(row => String(row?.item?.comment || "").trim()).filter(Boolean))].join("\n");
            group.images = [...new Set(group.rows.map(row => String(row?.item?.image || "").trim()).filter(Boolean))];
            group.image = group.images[0] || "";
        }
        return groups;
    }

    function isoWeekInputForDateKey(value) {
        const source = /^\d{4}-\d{2}-\d{2}$/.test(String(value || "")) ? String(value) : transactionDateKey(new Date());
        const parts = source.split("-").map(Number);
        const d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
        d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
        const year = d.getUTCFullYear();
        const yearStart = new Date(Date.UTC(year, 0, 1));
        const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
        return `${year}-W${String(week).padStart(2, "0")}`;
    }

    function isoWeekDateRange(value) {
        const match = String(value || "").match(/^(\d{4})-W(\d{2})$/);
        if (!match) return null;
        const year = Number(match[1]), week = Number(match[2]);
        if (week < 1 || week > 53) return null;
        const jan4 = new Date(Date.UTC(year, 0, 4));
        const mondayOffset = (jan4.getUTCDay() + 6) % 7;
        const monday = new Date(Date.UTC(year, 0, 4 - mondayOffset + (week - 1) * 7));
        // Reject an invalid week 53 instead of silently rolling into another ISO year.
        const check = isoWeekInputForDateKey(monday.toISOString().slice(0, 10));
        if (check !== `${year}-W${String(week).padStart(2, "0")}`) return null;
        const sunday = new Date(monday.getTime() + 6 * 86400000);
        return {from:monday.toISOString().slice(0, 10), to:sunday.toISOString().slice(0, 10), value:`${year}-W${String(week).padStart(2, "0")}`};
    }

    const clientLedgerExportPhotoFolderPromises = new Map();
    const clientLedgerExportPhotoFolderIds = new Map();
    const clientLedgerExportPhotoFilesByFolder = new Map();

    const CLIENT_LEDGER_PHOTO_ARCHIVE_ROOT_ID =
        CONFIG.DRIVE_CLIENT_LEDGER_PHOTOS_PARENT_FOLDER_ID || "1hs6fIiuxvqmj7VL4dsqowu6EANgauoik";

    function clientLedgerPhotoFolderName(client) {
        const displayClient = String(client || "Unknown Client").trim().replace(/\s+/g, " ") || "Unknown Client";
        return `${displayClient.toLocaleUpperCase()} MOVEMENTS`;
    }

    async function getOrCreateClientLedgerPhotoFolder(client) {
        const displayClient = String(client || "Unknown Client").trim().replace(/\s+/g, " ") || "Unknown Client";
        const clientKey = displayClient.toLocaleLowerCase();
        if (clientLedgerExportPhotoFolderIds.has(clientKey)) return clientLedgerExportPhotoFolderIds.get(clientKey);
        if (clientLedgerExportPhotoFolderPromises.has(clientKey)) return clientLedgerExportPhotoFolderPromises.get(clientKey);

        const pending = (async () => {
            const parent = CLIENT_LEDGER_PHOTO_ARCHIVE_ROOT_ID;
            const name = clientLedgerPhotoFolderName(displayClient);
            const query = [
                `name = '${escapeDriveQuery(name)}'`,
                `mimeType = 'application/vnd.google-apps.folder'`,
                `'${parent}' in parents`,
                `trashed = false`
            ].join(" and ");
            const params = new URLSearchParams({
                q: query,
                pageSize: "10",
                fields: "files(id,name,mimeType)",
                supportsAllDrives: "true",
                includeItemsFromAllDrives: "true"
            });
            const found = await fetchJson(`${DRIVE_API}?${params.toString()}`, {headers:authHeaders()});
            let id = (found.files || []).find(file => file.mimeType === "application/vnd.google-apps.folder")?.id || "";
            if (!id) {
                const created = await fetchJson(`${DRIVE_API}?supportsAllDrives=true&fields=id,name,mimeType`, {
                    method: "POST",
                    headers: {...authHeaders(), "Content-Type": "application/json"},
                    body: JSON.stringify({
                        name,
                        mimeType: "application/vnd.google-apps.folder",
                        parents: [parent],
                        description: `Client ledger report photos for ${displayClient}. Original transaction photos are not moved or deleted.`
                    })
                });
                id = created.id || "";
            }
            if (!id) throw new Error(`Google Drive did not return a folder ID for ${name}.`);
            clientLedgerExportPhotoFolderIds.set(clientKey, id);
            return id;
        })().catch(error => {
            clientLedgerExportPhotoFolderPromises.delete(clientKey);
            throw error;
        });
        clientLedgerExportPhotoFolderPromises.set(clientKey, pending);
        return pending;
    }

    async function loadClientLedgerPhotoFolderFiles(folderId) {
        if (clientLedgerExportPhotoFilesByFolder.has(folderId)) {
            return clientLedgerExportPhotoFilesByFolder.get(folderId);
        }
        const files = new Map();
        let pageToken = "";
        do {
            const params = new URLSearchParams({
                q: `'${folderId}' in parents and trashed = false`,
                pageSize: "1000",
                fields: "nextPageToken,files(id,name,webViewLink,mimeType)",
                supportsAllDrives: "true",
                includeItemsFromAllDrives: "true"
            });
            if (pageToken) params.set("pageToken", pageToken);
            const result = await fetchJson(`${DRIVE_API}?${params.toString()}`, {headers:authHeaders()});
            for (const file of (result.files || [])) {
                const match = String(file.name || "").match(/^ledger-photo-.+__([A-Za-z0-9_-]+)$/);
                if (match && file.id) files.set(match[1], file);
            }
            pageToken = result.nextPageToken || "";
        } while (pageToken);
        clientLedgerExportPhotoFilesByFolder.set(folderId, files);
        return files;
    }

    async function archiveClientLedgerPhotosForExport(data, includePhotos) {
        const rootUrl = `https://drive.google.com/drive/folders/${CLIENT_LEDGER_PHOTO_ARCHIVE_ROOT_ID}`;
        if (!includePhotos) return {data, copied:0, archived:0, folderUrl:rootUrl, warning:""};

        // Store a distinct copy for each client, even when one source photo was
        // attached to a batch containing multiple clients.
        const sourceByClientAndId = new Map();
        for (const row of (data?.transactions || [])) {
            const url = String(row?.item?.image || "").trim();
            const sourceId = FM_MEDIA?.driveFileIdFromLink?.(url) || "";
            if (!sourceId) continue;
            const client = String(row?.item?.client || data?.client || "Unknown Client").trim() || "Unknown Client";
            const clientKey = client.toLocaleLowerCase();
            const key = `${clientKey}\u0000${sourceId}`;
            if (!sourceByClientAndId.has(key)) {
                sourceByClientAndId.set(key, {key, sourceId, client, clientKey, url, row});
            }
        }
        if (!sourceByClientAndId.size) return {data, copied:0, archived:0, folderUrl:rootUrl, warning:""};

        const remappedUrls = new Map();
        let copied = 0;
        try {
            const clients = [...new Map([...sourceByClientAndId.values()].map(info => [info.clientKey, info.client])).entries()];
            const folderByClientKey = new Map();
            await Promise.all(clients.map(async ([clientKey, client]) => {
                folderByClientKey.set(clientKey, await getOrCreateClientLedgerPhotoFolder(client));
            }));

            const jobs = [];
            const filesByClient = new Map();
            for (const info of sourceByClientAndId.values()) {
                const folderId = folderByClientKey.get(info.clientKey);
                if (!filesByClient.has(info.clientKey)) {
                    filesByClient.set(info.clientKey, await loadClientLedgerPhotoFolderFiles(folderId));
                }
                const files = filesByClient.get(info.clientKey);
                const existing = files.get(info.sourceId);
                if (existing?.id) {
                    remappedUrls.set(info.key, existing.webViewLink || `https://drive.google.com/file/d/${existing.id}/view`);
                } else {
                    jobs.push({...info, folderId, files});
                }
            }

            // Keep Drive API traffic modest for long histories with many photos.
            for (let i = 0; i < jobs.length; i += 3) {
                await Promise.all(jobs.slice(i, i + 3).map(async info => {
                    try {
                        const safeClient = String(info.client || "Client").replace(/[^a-z0-9]+/gi,"-").replace(/^-+|-+$/g,"").slice(0,40) || "Client";
                        const safeDate = String(info.row?.date || "undated").replace(/[^0-9-]/g, "");
                        const safeTime = String(info.row?.time || "").replace(/[^0-9]/g, "");
                        const exportPhotoName = `ledger-photo-${safeClient}_${safeDate}_${safeTime}__${info.sourceId}`;
                        const result = await fetchJson(`${DRIVE_API}/${encodeURIComponent(info.sourceId)}/copy?supportsAllDrives=true&fields=id,name,webViewLink,mimeType`, {
                            method: "POST",
                            headers: {...authHeaders(), "Content-Type": "application/json"},
                            body: JSON.stringify({
                                name: exportPhotoName,
                                parents: [info.folderId],
                                description: `Photo copy used in a client ledger report for ${info.client}.`,
                                appProperties: {ledgerExportSourceId: info.sourceId, ledgerExportClient: info.clientKey}
                            })
                        });
                        if (result?.id) {
                            const file = {...result, name:exportPhotoName};
                            info.files.set(info.sourceId, file);
                            remappedUrls.set(info.key, file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`);
                            copied += 1;
                        }
                    } catch (error) {
                        console.warn(`Could not copy ledger photo ${info.sourceId} into ${clientLedgerPhotoFolderName(info.client)}:`, error);
                    }
                }));
            }

            const remappedData = {
                ...data,
                transactions:(data.transactions || []).map(row => {
                    const original = String(row?.item?.image || "").trim();
                    const sourceId = FM_MEDIA?.driveFileIdFromLink?.(original) || "";
                    const client = String(row?.item?.client || data?.client || "Unknown Client").trim() || "Unknown Client";
                    const key = `${client.toLocaleLowerCase()}\u0000${sourceId}`;
                    const archivedUrl = sourceId ? remappedUrls.get(key) : "";
                    return archivedUrl ? {...row, item:{...row.item, image:archivedUrl}} : row;
                })
            };
            const unresolved = sourceByClientAndId.size - remappedUrls.size;
            return {
                data:remappedData,
                copied,
                archived:remappedUrls.size,
                folderUrl:rootUrl,
                warning:unresolved ? `${unresolved} photo copy/copies could not be archived; original photo links were kept.` : ""
            };
        } catch (error) {
            console.warn("Client ledger photo archive unavailable; retaining source photo links:", error);
            return {data, copied:0, archived:0, folderUrl:rootUrl, warning:"The client photo archive folder could not be accessed; original photo links were kept."};
        }
    }

    async function compactLedgerImageForExport(loaded, maxDimension = 640, quality = 0.68) {
        if (!loaded?.dataUrl || typeof document === "undefined") return loaded;
        try {
            const image = new Image();
            image.decoding = "async";
            image.src = loaded.dataUrl;
            await new Promise((resolve, reject) => { image.onload = resolve; image.onerror = reject; });
            const sourceWidth = image.naturalWidth || image.width || 1;
            const sourceHeight = image.naturalHeight || image.height || 1;
            const scale = Math.min(1, maxDimension / Math.max(sourceWidth, sourceHeight));
            const width = Math.max(1, Math.round(sourceWidth * scale));
            const height = Math.max(1, Math.round(sourceHeight * scale));
            const canvas = document.createElement("canvas");
            canvas.width = width; canvas.height = height;
            const context = canvas.getContext("2d", {alpha:false});
            if (!context) return loaded;
            context.drawImage(image, 0, 0, width, height);
            const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/jpeg", quality));
            if (!blob) return loaded;
            const dataUrl = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(String(reader.result || ""));
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
            return {dataUrl, extension:"jpeg"};
        } catch (error) {
            console.warn("Could not compress a ledger photo for export:", error);
            return loaded;
        }
    }

    function styleWorksheetHeader(row) {
        row.font = {bold:true, color:{argb:"FFFFFFFF"}};
        row.fill = {type:"pattern", pattern:"solid", fgColor:{argb:"FF211B32"}};
        row.alignment = {vertical:"middle", wrapText:true};
        row.height = 30;
        row.eachCell(cell => { cell.border = {bottom:{style:"thin", color:{argb:"FF4A4160"}}}; });
    }
    function styleWorksheetTotals(row) {
        row.font = {bold:true, color:{argb:"FFFFFFFF"}};
        row.fill = {type:"pattern", pattern:"solid", fgColor:{argb:"FF3A3152"}};
        row.alignment = {vertical:"middle", wrapText:true};
        row.height = 26;
    }
    async function addLedgerPhoto(workbook, worksheet, rowNumber, columnNumber, item, includePhotos, imageIdCache = new Map()) {
        if (!includePhotos || !item?.image) return;
        try {
            const url = String(item.image).trim();
            let imageId = imageIdCache.get(url);
            if (imageId == null) {
                const loaded = await FM_MEDIA?.loadImageDataUrl?.(url, state.accessToken);
                if (!loaded?.dataUrl) return;
                const compact = await compactLedgerImageForExport(loaded, 640, 0.68);
                imageId = workbook.addImage({base64:compact.dataUrl, extension:compact.extension || "jpeg"});
                imageIdCache.set(url, imageId);
            }
            worksheet.addImage(imageId, {tl:{col:columnNumber-1,row:rowNumber-1}, ext:{width:84,height:52}});
            worksheet.getRow(rowNumber).height = Math.max(44, worksheet.getRow(rowNumber).height || 15);
        } catch(error) { console.warn("Client ledger photo embed failed:", error); }
    }
    function applyLedgerCellBorders(cell, color="FF2B2B2B") {
        cell.border = {
            top:{style:"thin",color:{argb:color}},
            left:{style:"thin",color:{argb:color}},
            bottom:{style:"thin",color:{argb:color}},
            right:{style:"thin",color:{argb:color}}
        };
    }

    function mergeLedgerGroup(worksheet, startRow, endRow, columns) {
        if (endRow <= startRow) return;
        for (const column of columns) {
            const range = `${column}${startRow}:${column}${endRow}`;
            worksheet.mergeCells(range);
            const cell = worksheet.getCell(`${column}${startRow}`);
            cell.alignment = {vertical:"middle", horizontal:column === "D" ? "left" : "left", wrapText:true};
        }
    }

    function ledgerMainHeaders() {
        return ["Client", "Date", "Time", "Asset type", "Quantity sent", "Quantity received", "Comment", "Picture"];
    }

    function clientReportPhotoCell(item) {
        const url = String(item?.image || "").trim();
        if (!url) return "";
        // Export links point to an archived Drive copy when archiving was available;
        // source files are never moved or have their sharing settings changed.
        return `=HYPERLINK("${url.replace(/"/g, '""')}","Open photo")`;
    }

    function legacyClientReportPhotoMatchKey(client, date, time, asset, direction, quantity, comment) {
        const normalize = value => String(value ?? "").trim().toLowerCase();
        const normalizedTime = normalize(time).slice(0, 8);
        return [normalize(client), normalize(date), normalizedTime, normalize(asset), normalize(direction), Number(quantity) || 0, normalize(comment)].join("\u001f");
    }

    async function backfillLegacyClientReportPhotoLinks(spreadsheetId, rows) {
        const sourcesByKey = new Map();
        for (const item of getDashboardTransactions()) {
            if (!item?.image) continue;
            const {date, time} = exportDateTime(item);
            const direction = clientLedgerDirection(item.movement);
            if (!["Sent", "Received"].includes(direction)) continue;
            const key = legacyClientReportPhotoMatchKey(item.client, date, time, item.asset, direction, item.quantity, item.comment);
            if (!sourcesByKey.has(key)) sourcesByKey.set(key, []);
            sourcesByKey.get(key).push(item);
        }

        const pictureValues = rows.map(row => [String(row?.[9] ?? "")]);
        let changed = false;
        rows.forEach((row, index) => {
            const picture = String(row?.[9] ?? "").trim().toLowerCase();
            if (!["photo", "open photo"].includes(picture)) return;
            const sent = Number(String(row?.[4] ?? "0").replace(/,/g, "")) || 0;
            const received = Number(String(row?.[5] ?? "0").replace(/,/g, "")) || 0;
            const direction = sent ? "Sent" : received ? "Received" : "";
            const quantity = sent || received;
            if (!direction || !quantity) return;
            const key = legacyClientReportPhotoMatchKey(row?.[0], row?.[1], row?.[2], row?.[3], direction, quantity, row?.[8]);
            const queue = sourcesByKey.get(key);
            const matchingItem = queue?.shift();
            if (!matchingItem?.image) return;
            pictureValues[index] = [clientReportPhotoCell(matchingItem)];
            changed = true;
        });

        if (changed && rows.length) {
            const range = `${quoteSheetName("Client Ledger")}!J1:J${rows.length}`;
            await sheetsPut(`/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {
                range, majorDimension: "ROWS", values: pictureValues
            });
        }
    }

    function exportDateSummaryRows(rows) {
        const received = new Map(), sent = new Map();
        rows.forEach(row => {
            const asset = row.asset;
            const qty = clientLedgerQuantity(row.item);
            const direction = clientLedgerDirection(row.item.movement, row.item.quantity);
            if (direction === "Received") received.set(asset, (received.get(asset) || 0) + qty);
            else if (direction === "Sent") sent.set(asset, (sent.get(asset) || 0) + qty);
        });
        const keys = [...new Set([...received.keys(), ...sent.keys()])].sort((a,b) => a.localeCompare(b));
        return keys.map(asset => ({
            asset,
            receivedQty: received.get(asset) || 0,
            sentQty: sent.get(asset) || 0
        })).filter(row => row.receivedQty || row.sentQty);
    }

    function buildLedgerRowsForDate(client, dateGroup) {
        const rows = [];
        rows.push(["Date", dateGroup.date, "", "", "", "", "", ""]);
        rows.push(ledgerMainHeaders());
        const timestampGroups = ledgerTimestampGroups(dateGroup.rows);
        const groupByRow = new Map();
        timestampGroups.forEach(group => {
            for (let i = group.startIndex; i < group.endIndex; i += 1) groupByRow.set(i, group);
        });
        dateGroup.rows.forEach((row, index) => {
            const direction = clientLedgerDirection(row.item.movement, row.item.quantity);
            const qty = clientLedgerQuantity(row.item);
            const group = groupByRow.get(index);
            const isGroupStart = group?.startIndex === index;
            rows.push([
                index === 0 ? client : "",
                index === 0 ? row.date : "",
                isGroupStart ? row.time : "",
                row.asset,
                direction === "Sent" ? qty : 0,
                direction === "Received" ? qty : 0,
                isGroupStart ? group.comment : "",
                isGroupStart ? clientReportPhotoCell({image:group.image}) : ""
            ]);
        });
        for (const summary of exportDateSummaryRows(dateGroup.rows)) {
            rows.push([
                summary.receivedQty ? "Received" : "",
                summary.receivedQty ? summary.asset : "",
                summary.receivedQty || "",
                "",
                summary.sentQty ? "Sent" : "",
                summary.sentQty ? summary.asset : "",
                summary.sentQty || "",
                ""
            ]);
        }
        rows.push(new Array(8).fill(""));
        return rows;
    }

    async function ensureClientReportSpreadsheet(client) {
        await ensureWriteAccess();
        const rows = state.clientListRows || [];
        let {header, clientIndex, reportIndex} = findClientListColumns(rows);
        if (clientIndex < 0) throw new Error(`The "${CONFIG.CLIENT_LIST_SHEET_NAME}" sheet does not contain a client name column.`);
        const clientRow = findClientListRowNumber(rows, client, clientIndex);
        if (clientRow < 0) throw new Error(`Client "${client}" was not found in the Client List.`);

        if (reportIndex < 0) {
            reportIndex = header.length;
            header = header.concat(["Reports"]);
            const headerRange = `${quoteSheetName(CONFIG.CLIENT_LIST_SHEET_NAME)}!A1:${columnLetter(header.length)}1`;
            await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(headerRange)}?valueInputOption=USER_ENTERED`, {
                range: headerRange, majorDimension: "ROWS", values: [header]
            });
            rows[0] = header;
        }

        let existingValue = String(rows[clientRow - 1]?.[reportIndex] ?? "").trim();
        let spreadsheetId = spreadsheetIdFromLink(existingValue);
        if (spreadsheetId) {
            try {
                const meta = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}`);
                const sheets = meta.sheets || [];
                let reportSheet = sheets.find(sheet => String(sheet.properties?.title || "") === "Client Ledger");
                if (!reportSheet && sheets.length) {
                    reportSheet = sheets[0];
                    await sheetsPost(`/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {requests: [{
                        updateSheetProperties: {
                            properties: {sheetId: reportSheet.properties.sheetId, title: "Client Ledger"},
                            fields: "title"
                        }
                    }]});
                }
                return {spreadsheetId, spreadsheetUrl: clientReportUrl(spreadsheetId), sheet: reportSheet};
            } catch (error) {
                console.warn("Existing client report could not be opened; a new report will be created:", error);
                spreadsheetId = "";
            }
        }

        const created = await fetchJson(SHEETS_API, {
            method: "POST",
            headers: {...authHeaders(), "Content-Type": "application/json"},
            body: JSON.stringify({properties: {title: `${client} Asset Movements`}})
        });
        spreadsheetId = String(created.spreadsheetId || "").trim();
        if (!spreadsheetId && created.spreadsheetUrl) {
            spreadsheetId = spreadsheetIdFromLink(created.spreadsheetUrl);
        }
        if (!spreadsheetId) {
            throw new Error("Google created a client report but did not return its spreadsheet ID. Check Google Drive for the new report and retry after reconnecting Sheets access.");
        }

        // The create endpoint can occasionally omit `sheets` in its response even though
        // the spreadsheet exists. Re-read metadata by ID before declaring the export failed.
        let metadata = created;
        let reportSheet = (metadata.sheets || []).find(item => String(item.properties?.title || "") === "Client Ledger")
            || (metadata.sheets || []).find(item => item?.properties?.sheetId != null)
            || null;
        if (!reportSheet?.properties?.sheetId) {
            try {
                metadata = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}?includeGridData=false`);
                reportSheet = (metadata.sheets || []).find(item => String(item.properties?.title || "") === "Client Ledger")
                    || (metadata.sheets || []).find(item => item?.properties?.sheetId != null)
                    || null;
            } catch (metadataError) {
                console.warn("Could not immediately reload new client report metadata:", metadataError);
            }
        }
        if (!reportSheet?.properties?.sheetId) {
            try {
                const added = await sheetsPost(`/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
                    requests: [{addSheet: {properties: {title: "Client Ledger"}}}]
                });
                reportSheet = added.replies?.find(reply => reply.addSheet)?.addSheet || null;
            } catch (addSheetError) {
                console.warn("Could not add a report sheet; will retry metadata:", addSheetError);
            }
        }
        if (!reportSheet?.properties?.sheetId) {
            metadata = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}?includeGridData=false`);
            reportSheet = (metadata.sheets || []).find(item => String(item.properties?.title || "") === "Client Ledger")
                || (metadata.sheets || []).find(item => item?.properties?.sheetId != null)
                || null;
        }
        if (!reportSheet?.properties?.sheetId) {
            throw new Error(`The client report was created (${spreadsheetId}), but Google still did not return a writable sheet ID. The file is not lost; open it in Drive and retry the export.`);
        }
        if (String(reportSheet.properties.title || "") !== "Client Ledger") {
            await sheetsPost(`/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {requests: [{
                updateSheetProperties: {
                    properties: {sheetId: reportSheet.properties.sheetId, title: "Client Ledger"},
                    fields: "title"
                }
            }]});
            reportSheet.properties.title = "Client Ledger";
        }

        const reportUrl = clientReportUrl(spreadsheetId);
        const cellRange = `${quoteSheetName(CONFIG.CLIENT_LIST_SHEET_NAME)}!${columnLetter(reportIndex + 1)}${clientRow}`;
        await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(cellRange)}?valueInputOption=RAW`, {
            range: cellRange, majorDimension: "ROWS", values: [[reportUrl]]
        });
        const targetRow = rows[clientRow - 1] || [];
        while (targetRow.length <= reportIndex) targetRow.push("");
        targetRow[reportIndex] = reportUrl;
        rows[clientRow - 1] = targetRow;
        state.clientListRows = rows;
        return {spreadsheetId, spreadsheetUrl: reportUrl, sheet: reportSheet};
    }

    async function ensureClientReportLayout(report, client, data) {
        const sheet = await getClientReportSheet(report.spreadsheetId);
        const sheetId = sheet.properties.sheetId;
        let existing = await getValues(report.spreadsheetId, "Client Ledger");
        // Migrate reports created by the old ten-column export: delete the obsolete
        // Difference for date and All-time difference columns. This moves Comment
        // and Picture back into columns G and H for the already-exported history too.
        const hasLegacyDifferenceColumns = existing.some(row =>
            String(row?.[6] || "").trim().toLowerCase() === "difference for date" &&
            String(row?.[7] || "").trim().toLowerCase() === "all-time difference"
        );
        if (hasLegacyDifferenceColumns) {
            // Old client sheets stored only the word "Photo" in the picture cell.
            // Reconnect those cells to matching source Drive files where possible before
            // deleting the old difference columns (which shifts Picture into column H).
            try {
                await backfillLegacyClientReportPhotoLinks(report.spreadsheetId, existing);
            } catch (photoMigrationError) {
                console.warn("Could not backfill older client report photo links:", photoMigrationError);
            }
            await sheetsPost(`/${encodeURIComponent(report.spreadsheetId)}:batchUpdate`, {requests: [{
                deleteDimension: {range: {sheetId, dimension: "COLUMNS", startIndex: 6, endIndex: 8}}
            }]});
            existing = await getValues(report.spreadsheetId, "Client Ledger");
        }
        const oldHistoryStart = reportHistoryStartRow(existing);
        const desiredHistoryStart = 6 + Math.max(1, data.assets.length);
        const requests = [];

        if (oldHistoryStart && oldHistoryStart < desiredHistoryStart) {
            requests.push({insertDimension: {range: {sheetId, dimension: "ROWS", startIndex: oldHistoryStart - 1, endIndex: desiredHistoryStart - 1}, inheritFromBefore: false}});
        } else if (oldHistoryStart && oldHistoryStart > desiredHistoryStart) {
            requests.push({deleteDimension: {range: {sheetId, dimension: "ROWS", startIndex: desiredHistoryStart - 1, endIndex: oldHistoryStart - 1}}});
        }
        if (requests.length) await sheetsPost(`/${encodeURIComponent(report.spreadsheetId)}:batchUpdate`, {requests});

        const clearEnd = Math.max(4 + data.assets.length, 4);
        await sheetsPost(`/${encodeURIComponent(report.spreadsheetId)}:batchUpdate`, {requests: [{
            unmergeCells: {range: {sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 2}}
        }]});
        const clearRange = `${quoteSheetName("Client Ledger")}!A1:H${clearEnd}`;
        await sheetsPost(`/${encodeURIComponent(report.spreadsheetId)}/values/${encodeURIComponent(clearRange)}:clear`, {});

        const summaryRows = [
            [`${client} Movements`, ""],
            ["", ""],
            [`All-time balance with ${client}`, ""],
            ["Asset Type", "Net Balance"],
            ...data.assets.map(item => [item.asset, item.balance])
        ];
        const summaryRange = `${quoteSheetName("Client Ledger")}!A1:B${summaryRows.length}`;
        await sheetsPut(`/${encodeURIComponent(report.spreadsheetId)}/values/${encodeURIComponent(summaryRange)}?valueInputOption=USER_ENTERED`, {
            range: summaryRange, majorDimension: "ROWS", values: summaryRows
        });

        const formatting = [
            {mergeCells: {range: {sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 2}, mergeType: "MERGE_ALL"}},
            {repeatCell: {range: {sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8}, cell: {userEnteredFormat: {textFormat: {bold: true, fontSize: 16}, verticalAlignment: "MIDDLE"}}, fields: "userEnteredFormat(textFormat,verticalAlignment)"}},
            {repeatCell: {range: {sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 2}, cell: {userEnteredFormat: {textFormat: {bold: true}, horizontalAlignment: "CENTER", borders: {top: {style: "SOLID"}, bottom: {style: "SOLID"}, left: {style: "SOLID"}, right: {style: "SOLID"}}}}, fields: "userEnteredFormat(textFormat,horizontalAlignment,borders)"}},
            {repeatCell: {range: {sheetId, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}, cell: {userEnteredFormat: {textFormat: {bold: true}, borders: {bottom: {style: "SOLID"}}}}, fields: "userEnteredFormat(textFormat,borders)"}},
            {repeatCell: {range: {sheetId, startRowIndex: 4, endRowIndex: 4 + data.assets.length, startColumnIndex: 0, endColumnIndex: 2}, cell: {userEnteredFormat: {borders: {bottom: {style: "SOLID"}}}}, fields: "userEnteredFormat(borders)"}},
            {repeatCell: {range: {sheetId, startRowIndex: 4, endRowIndex: 4 + data.assets.length, startColumnIndex: 1, endColumnIndex: 2}, cell: {userEnteredFormat: {numberFormat: {type: "NUMBER", pattern: "#,##0"}, horizontalAlignment: "RIGHT"}}, fields: "userEnteredFormat(numberFormat,horizontalAlignment)"}},
            {updateDimensionProperties: {range: {sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1}, properties: {pixelSize: 150}, fields: "pixelSize"}},
            {updateDimensionProperties: {range: {sheetId, dimension: "COLUMNS", startIndex: 1, endIndex: 2}, properties: {pixelSize: 105}, fields: "pixelSize"}},
            {updateDimensionProperties: {range: {sheetId, dimension: "COLUMNS", startIndex: 2, endIndex: 3}, properties: {pixelSize: 85}, fields: "pixelSize"}},
            {updateDimensionProperties: {range: {sheetId, dimension: "COLUMNS", startIndex: 3, endIndex: 4}, properties: {pixelSize: 165}, fields: "pixelSize"}},
            {updateDimensionProperties: {range: {sheetId, dimension: "COLUMNS", startIndex: 4, endIndex: 6}, properties: {pixelSize: 115}, fields: "pixelSize"}},
            {updateDimensionProperties: {range: {sheetId, dimension: "COLUMNS", startIndex: 6, endIndex: 7}, properties: {pixelSize: 245}, fields: "pixelSize"}},
            {updateDimensionProperties: {range: {sheetId, dimension: "COLUMNS", startIndex: 7, endIndex: 8}, properties: {pixelSize: 135}, fields: "pixelSize"}}
        ];
        await sheetsPost(`/${encodeURIComponent(report.spreadsheetId)}:batchUpdate`, {requests: formatting});
        return {sheetId, historyStart: desiredHistoryStart};
    }

    async function getClientReportSheet(spreadsheetId) {
        const meta = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}`);
        let sheet = (meta.sheets || []).find(item => String(item.properties?.title || "") === "Client Ledger");
        if (!sheet && meta.sheets?.length) {
            sheet = meta.sheets[0];
            await sheetsPost(`/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {requests: [{
                updateSheetProperties: {properties: {sheetId: sheet.properties.sheetId, title: "Client Ledger"}, fields: "title"}
            }]});
        }
        if (!sheet?.properties?.sheetId) throw new Error("The client report does not contain a writable sheet.");
        return sheet;
    }

    function reportHistoryStartRow(rows) {
        for (let i = 0; i < rows.length; i += 1) {
            const first = String(rows[i]?.[0] ?? "").trim().toLowerCase();
            const second = String(rows[i]?.[1] ?? "").trim();
            if (first === "date" && /^\d{4}-\d{2}-\d{2}$/.test(second)) return i + 1;
        }
        return 0;
    }

    function blockFormattingRequests(sheetId, startRow, dateGroups) {
        const requests = [];
        let rowCursor = startRow - 1;
        for (const dateGroup of dateGroups) {
            const dateRow = rowCursor;
            const headerRow = rowCursor + 1;
            const dataStart = rowCursor + 2;
            const dataEnd = dataStart + dateGroup.rows.length;
            const summaryCount = exportDateSummaryRows(dateGroup.rows).length;
            requests.push({repeatCell: {range: {sheetId, startRowIndex: dateRow, endRowIndex: dateRow + 1, startColumnIndex: 0, endColumnIndex: 2}, cell: {userEnteredFormat: {textFormat: {bold: true}}}, fields: "userEnteredFormat(textFormat)"}});
            requests.push({repeatCell: {range: {sheetId, startRowIndex: headerRow, endRowIndex: headerRow + 1, startColumnIndex: 0, endColumnIndex: 8}, cell: {userEnteredFormat: {backgroundColor: {red: 0.13, green: 0.09, blue: 0.23}, textFormat: {bold: true, foregroundColor: {red: 1, green: 1, blue: 1}}, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE", wrapStrategy: "WRAP", borders: {top: {style: "SOLID"}, bottom: {style: "SOLID"}, left: {style: "SOLID"}, right: {style: "SOLID"}}}}, fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy,borders)"}});
            if (dateGroup.rows.length) {
                requests.push({mergeCells: {range: {sheetId, startRowIndex: dataStart, endRowIndex: dataEnd, startColumnIndex: 0, endColumnIndex: 1}, mergeType: "MERGE_ALL"}});
                requests.push({mergeCells: {range: {sheetId, startRowIndex: dataStart, endRowIndex: dataEnd, startColumnIndex: 1, endColumnIndex: 2}, mergeType: "MERGE_ALL"}});
                requests.push({repeatCell: {range: {sheetId, startRowIndex: dataStart, endRowIndex: dataEnd, startColumnIndex: 0, endColumnIndex: 8}, cell: {userEnteredFormat: {verticalAlignment: "MIDDLE", wrapStrategy: "WRAP", borders: {top: {style: "SOLID"}, bottom: {style: "SOLID"}, left: {style: "SOLID"}, right: {style: "SOLID"}}}}, fields: "userEnteredFormat(verticalAlignment,wrapStrategy,borders)"}});
                requests.push({repeatCell: {range: {sheetId, startRowIndex: dataStart, endRowIndex: dataEnd, startColumnIndex: 4, endColumnIndex: 6}, cell: {userEnteredFormat: {numberFormat: {type: "NUMBER", pattern: "#,##0"}, horizontalAlignment: "RIGHT"}}, fields: "userEnteredFormat(numberFormat,horizontalAlignment)"}});
                // Merge time, comment and picture fields for each same-timestamp movement.
                for (const group of ledgerTimestampGroups(dateGroup.rows)) {
                    if (group.count <= 1) continue;
                    const start = dataStart + group.startIndex;
                    const end = dataStart + group.endIndex;
                    for (const column of [2, 6, 7]) {
                        requests.push({mergeCells: {range: {sheetId, startRowIndex:start, endRowIndex:end, startColumnIndex:column, endColumnIndex:column + 1}, mergeType:"MERGE_ALL"}});
                    }
                }
            }
            if (summaryCount) {
                requests.push({repeatCell: {range: {sheetId, startRowIndex: dataEnd, endRowIndex: dataEnd + summaryCount, startColumnIndex: 0, endColumnIndex: 8}, cell: {userEnteredFormat: {textFormat: {bold: true}}}, fields: "userEnteredFormat(textFormat)"}});
                requests.push({repeatCell: {range: {sheetId, startRowIndex: dataEnd, endRowIndex: dataEnd + summaryCount, startColumnIndex: 2, endColumnIndex: 3}, cell: {userEnteredFormat: {numberFormat: {type: "NUMBER", pattern: "#,##0"}, horizontalAlignment: "RIGHT"}}, fields: "userEnteredFormat(numberFormat,horizontalAlignment)"}});
                requests.push({repeatCell: {range: {sheetId, startRowIndex: dataEnd, endRowIndex: dataEnd + summaryCount, startColumnIndex: 6, endColumnIndex: 7}, cell: {userEnteredFormat: {numberFormat: {type: "NUMBER", pattern: "#,##0"}, horizontalAlignment: "RIGHT"}}, fields: "userEnteredFormat(numberFormat,horizontalAlignment)"}});
            }
            rowCursor = dataEnd + summaryCount + 1;
        }
        return requests;
    }

    async function writeClientReportExport(report, client, data) {
        const layout = await ensureClientReportLayout(report, client, data);
        const dateGroups = exportDateGroups(data);
        if (!dateGroups.length) throw new Error("No transactions matched the selected period and asset filters.");
        const rows = dateGroups.flatMap(group => buildLedgerRowsForDate(client, group));
        const startRow = layout.historyStart;
        const endRow = startRow + rows.length - 1;
        await sheetsPost(`/${encodeURIComponent(report.spreadsheetId)}:batchUpdate`, {requests: [{
            insertDimension: {range: {sheetId: layout.sheetId, dimension: "ROWS", startIndex: startRow - 1, endIndex: endRow}, inheritFromBefore: false}
        }]});
        const valueRange = `${quoteSheetName("Client Ledger")}!A${startRow}:H${endRow}`;
        await sheetsPut(`/${encodeURIComponent(report.spreadsheetId)}/values/${encodeURIComponent(valueRange)}?valueInputOption=USER_ENTERED`, {range: valueRange, majorDimension: "ROWS", values: rows});
        await sheetsPost(`/${encodeURIComponent(report.spreadsheetId)}:batchUpdate`, {requests: blockFormattingRequests(layout.sheetId, startRow, dateGroups)});
        return {...report, url: report.spreadsheetUrl, dateCount: dateGroups.length, transactionCount: data.transactions.length};
    }

    const CLIENT_LEDGER_PDF_SOURCES = {
        jsPDF: [
            "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js",
            "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
            "https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js"
        ],
        autoTable: [
            "https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js",
            "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js",
            "https://unpkg.com/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"
        ]
    };
    let pdfJsLoadPromise = null;
    function hasClientLedgerPdfSupport() {
        const jsPDF = window.jspdf?.jsPDF;
        return Boolean(jsPDF && (typeof jsPDF.API?.autoTable === "function" || typeof jsPDF.prototype?.autoTable === "function"));
    }
    function loadExternalScript(sources, isReady) {
        return new Promise((resolve, reject) => {
            if (isReady()) return resolve();
            let index = 0;
            const tryNext = () => {
                if (isReady()) return resolve();
                if (index >= sources.length) return reject(new Error("PDF library failed to load. Please check your internet connection or browser content blocker."));
                const script = document.createElement("script");
                script.src = sources[index++];
                script.async = true;
                script.onload = () => isReady() ? resolve() : tryNext();
                script.onerror = tryNext;
                document.head.appendChild(script);
            };
            tryNext();
        });
    }
    function ensurePdfJs() {
        if (hasClientLedgerPdfSupport()) return Promise.resolve(window.jspdf.jsPDF);
        if (pdfJsLoadPromise) return pdfJsLoadPromise;
        pdfJsLoadPromise = (async () => {
            const hasJsPdf = () => typeof window.jspdf?.jsPDF === "function";
            await loadExternalScript(CLIENT_LEDGER_PDF_SOURCES.jsPDF, hasJsPdf);
            await loadExternalScript(CLIENT_LEDGER_PDF_SOURCES.autoTable, hasClientLedgerPdfSupport);
            if (!hasClientLedgerPdfSupport()) throw new Error("PDF library failed to load. Please check your internet connection or browser content blocker.");
            return window.jspdf.jsPDF;
        })().catch(error => {
            pdfJsLoadPromise = null;
            throw error;
        });
        return pdfJsLoadPromise;
    }

    const MAX_EMBEDDED_CLIENT_LEDGER_PHOTOS = 24;

    async function loadClientLedgerPhotoData(data, includePhotos) {
        const photos = new Map();
        if (!includePhotos) return photos;
        // Most transaction rows in a batch share a timestamp/photo. Deduplicate before
        // loading and cap embedded previews so large PDFs remain responsive in browsers.
        const urls = [...new Set(ledgerTimestampGroups((data?.transactions || [])).map(group => group.image).filter(Boolean))]
            .slice(0, MAX_EMBEDDED_CLIENT_LEDGER_PHOTOS);
        let nextIndex = 0;
        const workerCount = Math.min(2, urls.length);
        const workers = Array.from({length: workerCount}, async () => {
            while (nextIndex < urls.length) {
                const url = urls[nextIndex++];
                try {
                    const loaded = await FM_MEDIA?.loadImageDataUrl?.(url, state.accessToken);
                    if (loaded?.dataUrl) photos.set(url, await compactLedgerImageForExport(loaded, 480, 0.56));
                } catch (error) {
                    console.warn("Client ledger photo could not be loaded:", error);
                }
            }
        });
        await Promise.all(workers);
        return photos;
    }

    function buildClientLedgerPrintHtml(client, data, periodLabel, includePhotos, photoData = new Map()) {
        const rowsHtml = exportDateGroups(data).map(dateGroup => {
            const body = ledgerTimestampGroups(dateGroup.rows).map(group => {
                const first = group.rows[0];
                const count = group.count;
                const photoDataItem = photoData.get(group.image);
                const photo = includePhotos && group.image
                    ? (photoDataItem
                        ? `<a href="${escapeAttr(group.image)}" target="_blank" rel="noopener" class="photo-cell"><img loading="lazy" decoding="async" src="${escapeAttr(photoDataItem.dataUrl)}" alt="Transaction photo"></a>`
                        : `<a href="${escapeAttr(group.image)}" target="_blank" rel="noopener">Open photo</a>`)
                    : "";
                return group.rows.map((row, rowIndex) => {
                    const direction = clientLedgerDirection(row.item.movement, row.item.quantity);
                    const qty = clientLedgerQuantity(row.item);
                    const merged = rowIndex === 0
                        ? `<td rowspan="${count}">${escapeHtml(client)}</td><td rowspan="${count}">${escapeHtml(row.date)}</td><td rowspan="${count}">${escapeHtml(row.time)}</td>`
                        : "";
                    const shared = rowIndex === 0
                        ? `<td rowspan="${count}">${escapeHtml(group.comment).replace(/\n/g, "<br>")}</td><td rowspan="${count}">${photo}</td>`
                        : "";
                    return `<tr>${merged}<td>${escapeHtml(row.asset)}</td><td>${direction === "Sent" ? formatNumber(qty) : "0"}</td><td>${direction === "Received" ? formatNumber(qty) : "0"}</td>${shared}</tr>`;
                }).join("");
            }).join("");
            const summaries = exportDateSummaryRows(dateGroup.rows);
            const received = summaries.filter(item => item.receivedQty).map(item => `<tr><td>${escapeHtml(item.asset)}</td><td>${formatNumber(item.receivedQty)}</td></tr>`).join("");
            const sent = summaries.filter(item => item.sentQty).map(item => `<tr><td>${escapeHtml(item.asset)}</td><td>${formatNumber(item.sentQty)}</td></tr>`).join("");
            return `<section class="date-block"><div class="date-title">Date: ${escapeHtml(dateGroup.date)}</div><table class="ledger"><thead><tr>${ledgerMainHeaders().map(h => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table><div class="summary-row"><table><thead><tr><th>Received on ${escapeHtml(dateGroup.date)}</th><th>Quantity</th></tr></thead><tbody>${received || '<tr><td colspan="2">0</td></tr>'}</tbody></table><table><thead><tr><th>Sent on ${escapeHtml(dateGroup.date)}</th><th>Quantity</th></tr></thead><tbody>${sent || '<tr><td colspan="2">0</td></tr>'}</tbody></table></div></section>`;
        }).join("");
        const assetRows = data.assets.map(item => `<tr><td>${escapeHtml(item.asset)}</td><td>${formatNumber(item.balance)}</td></tr>`).join("");
        return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(client)} Movements</title><style>
@page{size:A4 landscape;margin:9mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#111;font-size:9px;margin:0}.title{font-size:19px;font-weight:700;margin:0 0 4px}.period{font-size:10px;margin-bottom:9px}.balance-title{font-size:11px;font-weight:700;margin:4px 0 3px}.asset-summary{width:65mm;border-collapse:collapse;margin-bottom:9px}.asset-summary th,.asset-summary td,.ledger th,.ledger td,.summary-row th,.summary-row td{border:1px solid #777;padding:2px 3px}.asset-summary th,.ledger th,.summary-row th{background:#211b32;color:#fff;font-weight:700}.asset-summary td:last-child,.ledger td:nth-child(5),.ledger td:nth-child(6){text-align:right}.ledger{width:100%;border-collapse:collapse;table-layout:fixed}.ledger th,.ledger td{font-size:7px;vertical-align:middle;word-break:break-word}.ledger th:nth-child(1){width:10%}.ledger th:nth-child(2){width:8%}.ledger th:nth-child(3){width:6%}.ledger th:nth-child(4){width:12%}.ledger th:nth-child(5){width:8%}.ledger th:nth-child(6){width:9%}.ledger th:nth-child(7){width:34%}.ledger th:nth-child(8){width:13%}.date-block{margin-top:7px}.date-title{font-size:11px;font-weight:700;margin-bottom:3px}.summary-row{display:flex;gap:10px;margin-top:4px;break-inside:avoid}.summary-row table{width:62mm;border-collapse:collapse}.summary-row th,.summary-row td{font-size:8px}.photo-cell img{max-width:22mm;max-height:12mm;display:block;margin:auto}@media print{.no-print{display:none}}
</style></head><body><div class="title">${escapeHtml(client)} Movements</div><div class="period">Period: ${escapeHtml(periodLabel)}</div><div class="balance-title">All-time balance with ${escapeHtml(client)}</div><table class="asset-summary"><thead><tr><th>Asset Type</th><th>Net Balance</th></tr></thead><tbody>${assetRows}</tbody></table>${rowsHtml}</body></html>`;
    }

    async function printClientLedgerPdfFallback(client, data, periodLabel, includePhotos) {
        const printWindow = window.open("", "_blank", "width=1200,height=900");
        if (!printWindow) throw new Error("The PDF library could not be loaded and the browser blocked the print window. Please allow pop-ups for this site and try again.");
        printWindow.opener = null;
        const photoData = await loadClientLedgerPhotoData(data, includePhotos);
        printWindow.document.open();
        printWindow.document.write(buildClientLedgerPrintHtml(client, data, periodLabel, includePhotos, photoData));
        printWindow.document.close();
        await new Promise(resolve => {
            const waitForImages = () => {
                const images = Array.from(printWindow.document.images || []);
                if (!images.length) return setTimeout(resolve, 250);
                let remaining = images.length;
                const done = () => { if (--remaining <= 0) setTimeout(resolve, 250); };
                images.forEach(img => {
                    if (img.complete) done();
                    else { img.addEventListener("load", done, {once:true}); img.addEventListener("error", done, {once:true}); }
                });
            };
            setTimeout(waitForImages, 100);
        });
        printWindow.focus();
        printWindow.print();
    }

    async function buildClientLedgerPdf(client, data, periodLabel, includePhotos) {
        const jsPDF = await ensurePdfJs();
        const doc = new jsPDF({orientation:"landscape", unit:"mm", format:"a4", compress:true, precision:2});
        const margin = 9;
        const photoData = await loadClientLedgerPhotoData(data, includePhotos);
        doc.setFontSize(16); doc.setFont(undefined, "bold"); doc.text(`${client} Movements`, margin, 12);
        doc.setFontSize(8); doc.setFont(undefined, "normal"); doc.text(`Period: ${periodLabel}`, margin, 18);
        if (includePhotos) {
            const photoCount = new Set(ledgerTimestampGroups(data.transactions).map(group => group.image).filter(Boolean)).size;
            doc.setFontSize(6.5); doc.text(`Photos: compact previews for ${Math.min(photoCount, MAX_EMBEDDED_CLIENT_LEDGER_PHOTOS)} of ${photoCount} unique movements; other photo cells remain clickable links.`, margin, 22);
        }
        doc.setFontSize(10); doc.setFont(undefined, "bold"); doc.text(`All-time balance with ${client}`, margin, 27);
        doc.autoTable({startY:30, head:[["Asset Type","Net Balance"]], body:data.assets.map(item=>[item.asset,formatNumber(item.balance)]), theme:"grid", styles:{fontSize:8,cellPadding:1.6}, headStyles:{fillColor:[33,27,50],textColor:255}, columnStyles:{0:{cellWidth:46},1:{cellWidth:27,halign:"right"}}});
        let y = doc.lastAutoTable.finalY + 5;
        for (const dateGroup of exportDateGroups(data)) {
            if (y > 180) { doc.addPage(); y = 12; }
            doc.setFontSize(10); doc.setFont(undefined, "bold"); doc.text(`Date: ${dateGroup.date}`, margin, y); doc.setFont(undefined, "normal");
            const body = [];
            const photoByBodyRow = new Map();
            for (const group of ledgerTimestampGroups(dateGroup.rows)) {
                const first = group.rows[0];
                const firstBodyRow = body.length;
                const photo = photoData.get(group.image);
                const canEmbed = Boolean(includePhotos && group.image && photo?.dataUrl && ["jpeg", "png"].includes(String(photo.extension || "").toLowerCase()));
                const span = value => group.count > 1 ? {content:value, rowSpan:group.count, styles:{valign:"middle"}} : value;
                const firstQty = clientLedgerQuantity(first.item);
                const firstDirection = clientLedgerDirection(first.item.movement, first.item.quantity);
                body.push([
                    span(client), span(first.date), span(first.time), first.asset,
                    firstDirection === "Sent" ? formatNumber(firstQty) : "0",
                    firstDirection === "Received" ? formatNumber(firstQty) : "0",
                    span(group.comment),
                    span(includePhotos && group.image ? (canEmbed ? "" : "Open photo") : "")
                ]);
                if (includePhotos && group.image) photoByBodyRow.set(firstBodyRow, {url:group.image, image:photo, canEmbed});
                for (let index = 1; index < group.rows.length; index += 1) {
                    const row = group.rows[index];
                    const direction = clientLedgerDirection(row.item.movement, row.item.quantity);
                    const qty = clientLedgerQuantity(row.item);
                    // AutoTable supports ragged rows when earlier cells use rowSpan.
                    body.push([row.asset, direction === "Sent" ? formatNumber(qty) : "0", direction === "Received" ? formatNumber(qty) : "0"]);
                }
            }
            doc.autoTable({
                startY:y + 3,
                head:[ledgerMainHeaders()], body,
                theme:"grid", margin:{left:margin,right:margin}, rowPageBreak:"avoid",
                styles:{fontSize:6.1,cellPadding:1.0,overflow:"linebreak",valign:"middle",minCellHeight:5},
                headStyles:{fillColor:[33,27,50],textColor:255,fontSize:6.0},
                columnStyles:{0:{cellWidth:23},1:{cellWidth:22},2:{cellWidth:16},3:{cellWidth:32},4:{cellWidth:18,halign:"right"},5:{cellWidth:20,halign:"right"},6:{cellWidth:104},7:{cellWidth:27,halign:"center"}},
                didParseCell: hook => {
                    if (hook.section === "body" && hook.column.index === 7 && photoByBodyRow.has(hook.row.index)) hook.cell.styles.minCellHeight = 14;
                },
                didDrawCell: hook => {
                    if (hook.section !== "body" || hook.column.index !== 7 || !includePhotos) return;
                    const photoEntry = photoByBodyRow.get(hook.row.index);
                    if (!photoEntry) return;
                    if (photoEntry.canEmbed && photoEntry.image?.dataUrl) {
                        try {
                            const format = String(photoEntry.image.extension).toLowerCase() === "png" ? "PNG" : "JPEG";
                            const imageWidth = Math.min(19, hook.cell.width - 2);
                            const imageHeight = Math.min(11, hook.cell.height - 2);
                            doc.addImage(photoEntry.image.dataUrl, format, hook.cell.x + (hook.cell.width - imageWidth) / 2, hook.cell.y + (hook.cell.height - imageHeight) / 2, imageWidth, imageHeight, undefined, "FAST");
                        } catch (error) {
                            console.warn("Could not embed a compact client-ledger photo in the PDF:", error);
                        }
                    }
                    doc.link(hook.cell.x, hook.cell.y, hook.cell.width, hook.cell.height, {url:photoEntry.url});
                }
            });
            y = doc.lastAutoTable.finalY + 3;
            const summaries = exportDateSummaryRows(dateGroup.rows);
            if (y > 184) { doc.addPage(); y = 12; }
            const receivedBody = summaries.filter(item=>item.receivedQty).map(item=>[item.asset,formatNumber(item.receivedQty)]);
            const sentBody = summaries.filter(item=>item.sentQty).map(item=>[item.asset,formatNumber(item.sentQty)]);
            doc.autoTable({startY:y,margin:{left:margin,right:210},tableWidth:62,head:[[`Received on ${dateGroup.date}`,"Quantity"]],body:receivedBody,theme:"grid",styles:{fontSize:7,cellPadding:1.1},headStyles:{fillColor:[33,27,50],textColor:255}});
            const receivedFinalY = doc.lastAutoTable.finalY;
            doc.autoTable({startY:y,margin:{left:156,right:margin},tableWidth:62,head:[[`Sent on ${dateGroup.date}`,"Quantity"]],body:sentBody,theme:"grid",styles:{fontSize:7,cellPadding:1.1},headStyles:{fillColor:[33,27,50],textColor:255}});
            const sentFinalY = doc.lastAutoTable.finalY;
            y = Math.max(receivedFinalY,sentFinalY,y+11) + 5;
        }
        return doc;
    }

    async function buildClientWeeklySummaryPdf(client, data, weekRange) {
        const jsPDF = await ensurePdfJs();
        const doc = new jsPDF({orientation:"landscape", unit:"mm", format:"a4", compress:true, precision:2});
        const margin = 12;
        const daily = new Map();
        for (const row of (data?.dailyRows || [])) {
            const key = `${row.date}\u0000${row.asset}`;
            daily.set(key, {date:row.date,asset:row.asset,sent:Number(row.sent)||0,received:Number(row.received)||0,net:(Number(row.sent)||0)-(Number(row.received)||0)});
        }
        const assets = [...new Set([...(data.assets || []).map(item=>item.asset), ...[...daily.values()].map(row=>row.asset)])].sort((a,b)=>a.localeCompare(b));
        const dates = [];
        for (let time = Date.parse(`${weekRange.from}T00:00:00Z`); time <= Date.parse(`${weekRange.to}T00:00:00Z`); time += 86400000) dates.push(new Date(time).toISOString().slice(0,10));

        // Keep the weekly daily-summary table truly daily: exactly one row for each day,
        // with that day's per-asset figures listed compactly inside the Sent/Received/Net cells.
        const dailyBody = dates.map(date => {
            const dayRows = assets
                .map(asset => daily.get(`${date}\u0000${asset}`) || {date,asset,sent:0,received:0,net:0})
                .filter(row => row.sent || row.received);
            if (!dayRows.length) return [date, "No movements", "—", "—"];
            const sent = dayRows.filter(row=>row.sent).map(row=>`${row.asset}: ${formatNumber(row.sent)}`).join("\n");
            const received = dayRows.filter(row=>row.received).map(row=>`${row.asset}: ${formatNumber(row.received)}`).join("\n");
            const net = dayRows.filter(row=>row.net).map(row=>`${row.asset}: ${row.net > 0 ? "+" : "-"}${formatNumber(Math.abs(row.net))}`).join("\n");
            return [date, sent || "—", received || "—", net || "—"];
        });

        const weeklyByAsset = new Map(assets.map(asset=>[asset,{asset,sent:0,received:0,net:0}]));
        for (const row of daily.values()) {
            if (row.date < weekRange.from || row.date > weekRange.to) continue;
            const target = weeklyByAsset.get(row.asset);
            if (!target) continue;
            target.sent += row.sent; target.received += row.received; target.net += row.net;
        }
        const weeklyBody = [...weeklyByAsset.values()]
            .filter(row=>row.sent || row.received)
            .map(row=>[row.asset,formatNumber(row.sent),formatNumber(row.received),`${row.net > 0 ? "+" : row.net < 0 ? "-" : ""}${formatNumber(Math.abs(row.net))}`]);
        const summaryRows = [
            ["Period",`${weekRange.from} to ${weekRange.to}`],
            ["Report format","One row per day with asset breakdowns; individual transactions are not included."]
        ];
        doc.setFontSize(17); doc.setFont(undefined,"bold"); doc.text(`${client} — Weekly Movement Report`,margin,15);
        doc.setFontSize(9); doc.setFont(undefined,"normal"); doc.text(`Week ${weekRange.value} · ${weekRange.from} to ${weekRange.to}`,margin,22);
        doc.autoTable({startY:26,head:[["Report detail","Value"]],body:summaryRows,theme:"grid",styles:{fontSize:8,cellPadding:1.8},headStyles:{fillColor:[33,27,50],textColor:255},columnStyles:{0:{cellWidth:42},1:{cellWidth:150}}});
        let y = doc.lastAutoTable.finalY + 6;
        doc.setFontSize(11); doc.setFont(undefined,"bold"); doc.text("Daily totals — one row per day",margin,y);
        doc.autoTable({
            startY:y+2,
            head:[["Date","Quantity sent by asset","Quantity received by asset","Net movement by asset"]],
            body:dailyBody.length?dailyBody:[["—","No transactions","—","—"]],
            theme:"grid",
            styles:{fontSize:7.2,cellPadding:1.8,overflow:"linebreak",valign:"middle",minCellHeight:8},
            headStyles:{fillColor:[33,27,50],textColor:255},
            columnStyles:{0:{cellWidth:30},1:{cellWidth:75},2:{cellWidth:75},3:{cellWidth:75}},
            rowPageBreak:"avoid"
        });
        y = doc.lastAutoTable.finalY + 8;
        if (y > 175) { doc.addPage(); y = 16; }
        doc.setFontSize(11); doc.setFont(undefined,"bold"); doc.text("Totals for the week",margin,y);
        doc.autoTable({startY:y+2,head:[["Asset type","Quantity sent","Quantity received","Net movement"]],body:weeklyBody.length?weeklyBody:[["No movements","0","0","0"]],theme:"grid",styles:{fontSize:8,cellPadding:1.8},headStyles:{fillColor:[33,27,50],textColor:255},columnStyles:{0:{cellWidth:82},1:{cellWidth:35,halign:"right"},2:{cellWidth:38,halign:"right"},3:{cellWidth:35,halign:"right"}}});
        return doc;
    }

    async function runClientLedgerExport() {
        const client = $("client-details-title")?.textContent?.trim() || "Unknown client";
        const fromKey = $("client-ledger-export-from")?.value || "";
        const toKey = $("client-ledger-export-to")?.value || "";
        const assets = selectedClientExportAssets();
        const includePhotos = Boolean($("client-ledger-export-photos")?.checked);
        const status = $("client-ledger-export-status"), button = $("run-client-ledger-export");
        if (fromKey && toKey && fromKey > toKey) { if(status){status.textContent="From date must be on or before To date.";status.classList.add("error");} return; }
        if (!assets.length) { if(status){status.textContent="No asset types are available for this client.";status.classList.add("error");} return; }
        let data = makeClientLedgerData(client, getDashboardTransactions(), fromKey, toKey, assets);
        if (!data.transactions.length) { if(status){status.textContent="No client transactions match that period and asset selection.";status.classList.add("error");} return; }
        const periodLabel = fromKey || toKey ? `${fromKey || "start"} to ${toKey || "today"}` : "All available history";
        let photoArchive = {copied:0,archived:0,warning:""};
        try {
            if(button){button.disabled=true;button.textContent="Saving…";}
            if(status){status.textContent=includePhotos ? "Copying photos into client-specific MOVEMENTS folders in Drive and preparing the report…" : "Saving the selected transactions to the client Google Sheet…";status.classList.remove("error");}
            photoArchive = await archiveClientLedgerPhotosForExport(data, includePhotos);
            data = photoArchive.data;
            const report = await ensureClientReportSpreadsheet(client);
            await writeClientReportExport(report, client, data);
            const wb = await buildClientLedgerWorkbook(client, data, periodLabel, includePhotos);
            const buffer = await wb.xlsx.writeBuffer();
            const url = URL.createObjectURL(new Blob([buffer], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
            const safeClient = client.replace(/[^a-z0-9]+/gi,"-").replace(/^-+|-+$/g,"").toLowerCase() || "client";
            const datePart = toKey || fromKey || transactionDateKey(new Date());
            const a=document.createElement("a"); a.href=url; a.download=`${safeClient}-client-ledger-${datePart}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
            setTimeout(()=>URL.revokeObjectURL(url),3000);
            if(status)status.innerHTML=`Saved to the client sheet. <a href="${escapeAttr(report.url)}" target="_blank" rel="noopener">Open client report</a>. XLSX copy downloaded.${photoArchive.archived ? ` Filed ${photoArchive.archived} photo copy/copies in the <a href="${escapeAttr(photoArchive.folderUrl)}" target="_blank" rel="noopener">client movement photo archive</a>, inside each client's MOVEMENTS folder.` : ""}${photoArchive.warning ? ` ${escapeHtml(photoArchive.warning)}` : ""}`;
        } catch(error) {
            console.error("Client ledger export failed:",error);
            if(status){status.textContent=error?.message||"The client report could not be saved.";status.classList.add("error");}
        } finally { if(button){button.disabled=false;button.textContent="Save XLSX + Client Sheet";} }
    }

    async function runClientLedgerPdfExport() {
        const client = $("client-details-title")?.textContent?.trim() || "Unknown client";
        const fromKey = $("client-ledger-export-from")?.value || "";
        const toKey = $("client-ledger-export-to")?.value || "";
        const assets = selectedClientExportAssets();
        const includePhotos = Boolean($("client-ledger-export-photos")?.checked);
        const status = $("client-ledger-export-status"), button = $("run-client-ledger-pdf");
        if (fromKey && toKey && fromKey > toKey) { if(status){status.textContent="From date must be on or before To date.";status.classList.add("error");} return; }
        let data = makeClientLedgerData(client, getDashboardTransactions(), fromKey, toKey, assets);
        if (!data.transactions.length) { if(status){status.textContent="No client transactions match that period and asset selection.";status.classList.add("error");} return; }
        const periodLabel = fromKey || toKey ? `${fromKey || "start"} to ${toKey || "today"}` : "All available history";
        let photoArchive = {copied:0,archived:0,warning:""};
        try {
            if(button){button.disabled=true;button.textContent="Building PDF…";}
            if(status){status.textContent=includePhotos ? "Copying photos into client-specific MOVEMENTS folders in Drive and building a compact PDF…" : "Building the PDF…";status.classList.remove("error");}
            photoArchive = await archiveClientLedgerPhotosForExport(data, includePhotos);
            data = photoArchive.data;
            const doc = await buildClientLedgerPdf(client,data,periodLabel,includePhotos);
            const safeClient = client.replace(/[^a-z0-9]+/gi,"-").replace(/^-+|-+$/g,"").toLowerCase() || "client";
            const datePart = toKey || fromKey || transactionDateKey(new Date());
            doc.save(`${safeClient}-client-ledger-${datePart}.pdf`);
            if(status)status.innerHTML=`PDF exported.${photoArchive.archived ? ` Filed ${photoArchive.archived} photo copy/copies in the <a href="${escapeAttr(photoArchive.folderUrl)}" target="_blank" rel="noopener">client movement photo archive</a>, inside each client's MOVEMENTS folder.` : ""}${photoArchive.warning ? ` ${escapeHtml(photoArchive.warning)}` : ""}`;
        } catch(error) {
            console.error("Client ledger PDF export failed:",error);
            if (/^PDF library failed to load\./i.test(error?.message || "")) {
                try {
                    if(status){status.textContent="PDF library unavailable. Opening the A4 print-to-PDF version…";status.classList.remove("error");}
                    await printClientLedgerPdfFallback(client, data, periodLabel, includePhotos);
                    if(status)status.textContent="Print-to-PDF opened. Choose Save as PDF in the browser print dialog.";
                } catch (fallbackError) {
                    if(status){status.textContent=fallbackError?.message||"The PDF could not be generated.";status.classList.add("error");}
                }
            } else if(status){status.textContent=error?.message||"The PDF could not be generated.";status.classList.add("error");}
        } finally { if(button){button.disabled=false;button.textContent="Export PDF";} }
    }

    async function runClientLedgerWeeklyPdfExport() {
        const client = $("client-details-title")?.textContent?.trim() || "Unknown client";
        const weekValue = $("client-ledger-export-week")?.value || "";
        const assets = selectedClientExportAssets();
        const status = $("client-ledger-export-status"), button = $("run-client-ledger-week-pdf");
        const weekRange = isoWeekDateRange(weekValue);
        if (!weekRange) {
            if (status) { status.textContent = "Choose a valid week before exporting the weekly report."; status.classList.add("error"); }
            return;
        }
        if (!assets.length) {
            if (status) { status.textContent = "Select at least one asset type for the weekly report."; status.classList.add("error"); }
            return;
        }
        const data = makeClientLedgerData(client, getDashboardTransactions(), weekRange.from, weekRange.to, assets);
        if (!data.transactions.length) {
            if (status) { status.textContent = `No ${client} movements were found for ${weekRange.from} to ${weekRange.to}.`; status.classList.add("error"); }
            return;
        }
        try {
            if (button) { button.disabled = true; button.textContent = "Building…"; }
            if (status) { status.textContent = "Building a lightweight weekly totals report…"; status.classList.remove("error"); }
            const doc = await buildClientWeeklySummaryPdf(client, data, weekRange);
            const safeClient = client.replace(/[^a-z0-9]+/gi,"-").replace(/^-+|-+$/g,"").toLowerCase() || "client";
            doc.save(`${safeClient}-weekly-movement-report-${weekRange.value}.pdf`);
            if (status) status.textContent = `Weekly PDF exported for ${weekRange.from} to ${weekRange.to}. It contains daily totals and a final weekly total, not individual transactions.`;
        } catch (error) {
            console.error("Client weekly PDF export failed:", error);
            if (status) { status.textContent = error?.message || "The weekly PDF could not be generated."; status.classList.add("error"); }
        } finally {
            if (button) { button.disabled = false; button.textContent = "Export Weekly PDF"; }
        }
    }

    async function buildClientLedgerWorkbook(client, data, periodLabel, includePhotos) {
        const ExcelJS = await ensureExcelJS();
        const wb = new ExcelJS.Workbook();
        wb.creator = "Assets Inventory Dashboard"; wb.created = new Date(); wb.modified = new Date();
        const ledger = wb.addWorksheet("Client Ledger", {views:[{state:"frozen",ySplit:1,showGridLines:false}]});
        ledger.columns = [
            {header:"Client",key:"client",width:18},{header:"Date",key:"date",width:13},{header:"Time",key:"time",width:11},{header:"Asset type",key:"asset",width:22},
            {header:"Quantity sent",key:"sent",width:15},{header:"Quantity received",key:"received",width:19},{header:"Comment",key:"comment",width:34},{header:"Picture",key:"picture",width:18}
        ];
        ledger.mergeCells("A1:H1"); ledger.getCell("A1").value = `${client} Movements`; ledger.getCell("A1").font = {bold:true,size:16};
        ledger.mergeCells("A3:B3"); ledger.getCell("A3").value = `All-time balance with ${client}`; ledger.getCell("A3").font = {bold:true}; ledger.getCell("A3").alignment = {horizontal:"center"};
        ledger.getRow(4).values = ["Asset Type","Net Balance"]; styleWorksheetHeader(ledger.getRow(4));
        data.assets.forEach(item=>{const r=ledger.addRow([item.asset,item.balance]);r.getCell(2).numFmt="#,##0";});
        let rowCursor = 6 + data.assets.length;
        const firstHistoryHeader = rowCursor + 1;
        const photoImageIds = new Map();
        for(const dateGroup of exportDateGroups(data)){
            ledger.getCell(`A${rowCursor}`).value="Date"; ledger.getCell(`B${rowCursor}`).value=dateGroup.date; ledger.getCell(`A${rowCursor}`).font={bold:true}; ledger.getCell(`B${rowCursor}`).font={bold:true};
            const headerRow=ledger.getRow(rowCursor+1); headerRow.values=ledgerMainHeaders(); styleWorksheetHeader(headerRow);
            const dataStart=headerRow.number+1;
            const timestampGroups=ledgerTimestampGroups(dateGroup.rows);
            const timestampGroupByRow=new Map();
            timestampGroups.forEach(group=>{for(let i=group.startIndex;i<group.endIndex;i+=1)timestampGroupByRow.set(i,group);});
            for (let index=0; index<dateGroup.rows.length; index+=1) {
                const row=dateGroup.rows[index];
                const group=timestampGroupByRow.get(index),isGroupStart=group?.startIndex===index;
                const direction=clientLedgerDirection(row.item.movement,row.item.quantity),qty=clientLedgerQuantity(row.item);
                const photoUrl=isGroupStart?group.image:"";
                const photoCell=photoUrl?{text:"Open photo",hyperlink:photoUrl}:"";
                const r=ledger.addRow([index===0?client:"",index===0?row.date:"",isGroupStart?row.time:"",row.asset,direction==="Sent"?qty:0,direction==="Received"?qty:0,isGroupStart?group.comment:"",photoCell]);
                r.height=isGroupStart && photoUrl?44:26;r.alignment={vertical:"middle",wrapText:true}; for(let c=1;c<=8;c++)applyLedgerCellBorders(r.getCell(c)); [5,6].forEach(c=>r.getCell(c).numFmt="#,##0");
                if(photoUrl){r.getCell(8).font={color:{argb:"FF6D5AA8"},underline:true};}
                if(includePhotos && photoUrl) await addLedgerPhoto(wb,ledger,r.number,8,{image:photoUrl},true,photoImageIds);
            }
            const dataEnd=ledger.rowCount;
            if(dateGroup.rows.length){
                ledger.mergeCells(`A${dataStart}:A${dataEnd}`); ledger.mergeCells(`B${dataStart}:B${dataEnd}`);
                for(const group of timestampGroups){
                    if(group.count<=1)continue;
                    const firstRow=dataStart+group.startIndex,lastRow=dataStart+group.endIndex-1;
                    for(const column of ["C","G","H"])ledger.mergeCells(`${column}${firstRow}:${column}${lastRow}`);
                }
            }
            for(const summary of exportDateSummaryRows(dateGroup.rows)){
                const r=ledger.addRow([summary.receivedQty?"Received":"",summary.receivedQty?summary.asset:"",summary.receivedQty||"","",summary.sentQty?"Sent":"",summary.sentQty?summary.asset:"",summary.sentQty||"",""]);
                r.font={bold:true}; r.getCell(3).numFmt="#,##0"; r.getCell(7).numFmt="#,##0";
            }
            ledger.addRow([]); rowCursor=ledger.rowCount+1;
        }
        if (ledger.rowCount >= firstHistoryHeader) ledger.autoFilter={from:`A${firstHistoryHeader}`,to:`H${ledger.rowCount}`};
        ledger.pageSetup={orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0};
        const notes=wb.addWorksheet("Ledger Notes",{views:[{showGridLines:false}]}); notes.columns=[{header:"Field",key:"field",width:30},{header:"Meaning",key:"meaning",width:100}]; styleWorksheetHeader(notes.getRow(1));
        notes.addRows([["Client",client],["Period",periodLabel],["Asset selection",data.assets.map(item=>item.asset).join(", ")||"All asset types"],["Ordering","Newest export dates appear first; earlier history stays below."],["Pictures","Photo links use descriptive copies in each client's MOVEMENTS folder under the configured Drive photo archive root; XLSX embeds compact previews when available."]]);
        notes.eachRow((row,i)=>{if(i>1)row.alignment={vertical:"top",wrapText:true};});
        return wb;
    }

    function toggleTransactionSummary(button) {
        const container = button.closest(".transaction-summary");
        const body = container?.querySelector("[data-summary-body]");
        if (!container || !body) return;
        const expanded = button.getAttribute("aria-expanded") === "true";
        body.classList.toggle("hidden", expanded);
        button.setAttribute("aria-expanded", String(!expanded));
        button.textContent = expanded ? "Expand totals" : "Hide totals";
        container.classList.toggle("is-expanded", !expanded);
    }

    function populateDashboardFilters() {
        const filters = getDashboardFilters();
        const clients = [...new Set(getDashboardTransactions().map(t => String(t.client || "").trim()).filter(Boolean))]
            .sort((a,b) => a.localeCompare(b));
        const assets = [...new Set(getDashboardInventory().map(i => String(i.asset || "").trim()).filter(Boolean))]
            .sort((a,b) => a.localeCompare(b));

        const clientSelect = $("dashboard-filter-client");
        const clientAuditSelect = $("dashboard-filter-client-audit");
        const assetSelect = $("dashboard-filter-asset");
        const assetAuditSelect = $("dashboard-filter-asset-audit");

        const clientOptions = `<option value="">All clients</option>${clients.map(c => `<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join("")}`;
        const assetOptionsHtml = `<option value="">All asset types</option>${assets.map(a => `<option value="${escapeAttr(a)}">${escapeHtml(a)}</option>`).join("")}`;

        if (clientSelect) {
            clientSelect.innerHTML = clientOptions;
            clientSelect.value = clients.includes(filters.client) ? filters.client : "";
        }
        if (clientAuditSelect) {
            clientAuditSelect.innerHTML = clientOptions;
            clientAuditSelect.value = clients.includes(filters.client) ? filters.client : "";
        }
        if (assetSelect) {
            assetSelect.innerHTML = assetOptionsHtml;
            assetSelect.value = assets.includes(filters.asset) ? filters.asset : "";
        }
        if (assetAuditSelect) {
            assetAuditSelect.innerHTML = assetOptionsHtml;
            assetAuditSelect.value = assets.includes(filters.asset) ? filters.asset : "";
        }

        // The audit day/range controls are independent of the dashboard's date filter.
        syncAuditDateControls();
    }

    function clearDashboardFilters() {
        ["dashboard-filter-client", "dashboard-filter-date", "dashboard-filter-asset",
         "dashboard-filter-client-audit", "dashboard-filter-asset-audit"]
            .forEach(id => { if ($(id)) $(id).value = ""; });
        auditDateFilter.mode = "today";
        auditDateFilter.day = "";
        auditDateFilter.from = "";
        auditDateFilter.to = "";
        resetAuditPaging();
        renderClientCards();
        renderWarehouseAssets();
        renderAudit();
        renderFilteredLoadTables();
    }

    function renderFilteredLoadTables() {
        const filters = getDashboardFilters();
        const todayKey = transactionDateKey(new Date().toISOString());
        const dateOk = !filters.date || filters.date === todayKey;
        const clientOk = item => !filters.client || String(item.client || "").trim() === filters.client;
        const assetOk = item => !filters.asset || String(item.asset || "").trim() === filters.asset;

        const due = dateOk
            ? state.todayLoads.filter(item => clientOk(item) && assetOk(item))
            : [];
        const inbound = due.filter(item => item.direction === "Inbound");
        const plannedOutbound = due.filter(item => item.direction === "Outbound");
        const estimate = sheinEstimateEnabled && dateOk
            ? FM_DASHBOARD_FEATURES.estimateSheinOutbound(state.transactions)
            : { available: false };
        const canShowShein = !filters.client || filters.client.toLocaleLowerCase() === "shein domestic";
        const canShowRed = !filters.asset || filters.asset.toLocaleLowerCase() === "red sleeves";
        const outbound = canShowShein && canShowRed
            ? FM_DASHBOARD_FEATURES.applySheinEstimate(plannedOutbound, estimate, sheinEstimateEnabled)
            : plannedOutbound;
        const detail = $("shein-estimate-status");
        if (detail) {
            detail.textContent = !sheinEstimateEnabled ? "" : !dateOk ? "" : !canShowShein || !canShowRed ? ""
                : estimate.available
                    ? `SHEIN Domestic · ${formatNumber(estimate.quantity)} Red Sleeves received since last dispatch ${formatTimestamp(estimate.lastSent)}.`
                    : "Shein estimate unavailable: no previous Red Sleeves dispatch recorded.";
            detail.classList.toggle("hidden", !detail.textContent);
        }

        $("inbound-badge").textContent = inbound.length.toLocaleString();
        $("outbound-badge").textContent = outbound.length.toLocaleString();
        renderLoadTable($("inbound-body"), inbound);
        renderLoadTable($("outbound-body"), outbound);

        [document.querySelector("#inbound-body")?.closest(".load-section"),
         document.querySelector("#outbound-body")?.closest(".load-section")]
            .filter(Boolean)
            .forEach(section => {
                const el = section.querySelector(".filter-note");
                if (el) el.remove();
            });
    }


    function renderGlobalSearch() {
        const input = $("global-search");
        const container = $("global-search-results");
        if (!input || !container) return;
        const query = String(input.value || "").trim().toLocaleLowerCase();
        if (!query) {
            container.innerHTML = "";
            container.classList.add("hidden");
            state.globalSearchRows = [];
            return;
        }

        const includesQuery = value => String(value ?? "").toLocaleLowerCase().includes(query);
        const valuesIncludeQuery = values => (values || []).some(includesQuery);
        const transactions = getDashboardTransactions();
        const groups = FM_TRANSACTION_DISPLAY?.groupTransactions?.(transactions) || [];
        const matchingGroups = groups.filter(group => {
            const groupValues = [group.client, group.timestamp, group.vehicleReg, group.comment, ...(group.comments || [])];
            for (const item of (group.items || [])) {
                groupValues.push(item?.client, item?.timestamp, item?.movement, item?.asset, item?.quantity, item?.user, item?.comment, item?.image, item?.inventoryBalance);
            }
            return valuesIncludeQuery(groupValues);
        });
        state.globalSearchRows = matchingGroups;

        const inventoryMatches = getDashboardInventory().filter(item => valuesIncludeQuery([item.asset, item.balance]));
        const clientMatches = getDashboardClients().filter(client => includesQuery(client));
        const loadMatches = (state.todayLoads || []).filter(item => valuesIncludeQuery(Object.values(item || {})));
        const routineMatches = (state.routines || []).filter(item => valuesIncludeQuery(Object.values(item || {})));
        const alertMatches = (state.alerts || []).filter(item => valuesIncludeQuery(Object.values(item || {})));

        const sections = [];
        if (matchingGroups.length) {
            sections.push(`<section class="global-search-group"><div class="global-search-group-heading"><h3>Movements and audit records</h3><span>${matchingGroups.length.toLocaleString()} match${matchingGroups.length === 1 ? "" : "es"}</span></div><div class="transaction-card-list global-search-transactions">${matchingGroups.slice(0, 40).map((group, index) => renderTransactionCard(group, index, "data-global-info-index")).join("")}</div>${matchingGroups.length > 40 ? `<p class="global-search-more">Showing 40 of ${matchingGroups.length.toLocaleString()} movement matches. Refine your search to narrow the list.</p>` : ""}</section>`);
        }
        if (inventoryMatches.length) {
            sections.push(`<section class="global-search-group"><div class="global-search-group-heading"><h3>Warehouse assets</h3><span>${inventoryMatches.length.toLocaleString()} match${inventoryMatches.length === 1 ? "" : "es"}</span></div><div class="global-search-tiles">${inventoryMatches.slice(0, 20).map(item => `<button type="button" class="global-search-tile" data-global-asset="${escapeAttr(item.asset)}"><span class="eyebrow">WAREHOUSE ASSET</span><strong>${escapeHtml(item.asset)}</strong><span>Current balance: ${formatNumber(item.balance)}</span></button>`).join("")}</div></section>`);
        }
        if (clientMatches.length) {
            sections.push(`<section class="global-search-group"><div class="global-search-group-heading"><h3>Clients</h3><span>${clientMatches.length.toLocaleString()} match${clientMatches.length === 1 ? "" : "es"}</span></div><div class="global-search-tiles">${clientMatches.slice(0, 20).map(client => `<button type="button" class="global-search-tile" data-global-client="${escapeAttr(client)}"><span class="eyebrow">CLIENT LEDGER</span><strong>${escapeHtml(client)}</strong><span>Open client balances and movement history</span></button>`).join("")}</div></section>`);
        }
        if (loadMatches.length || routineMatches.length || alertMatches.length) {
            const related = [];
            loadMatches.slice(0, 15).forEach(item => related.push(`<article class="global-search-item"><span class="eyebrow">TODAY'S LOAD</span><strong>${escapeHtml(item.client || item.destination || "Scheduled load")}</strong><span>${escapeHtml([item.direction, item.asset, item.quantity, item.destination, item.plannedTime].filter(Boolean).join(" · "))}</span></article>`));
            routineMatches.slice(0, 15).forEach(item => related.push(`<article class="global-search-item"><span class="eyebrow">ROUTINE EXPECTATION</span><strong>${escapeHtml(item.client || item.destination || item.asset || "Routine movement")}</strong><span>${escapeHtml([item.direction, item.asset, item.quantity, item.destination, item.notes, item.frequency, item.weekday].filter(Boolean).join(" · "))}</span></article>`));
            alertMatches.slice(0, 15).forEach(item => related.push(`<article class="global-search-item"><span class="eyebrow">ALERT · ${escapeHtml(item.status || "STATUS UNKNOWN")}</span><strong>${escapeHtml(item.client || item.asset || "Operational alert")}</strong><span>${escapeHtml(item.message || item.destination || item.resolutionComment || item.type || "")}</span></article>`));
            sections.push(`<section class="global-search-group"><div class="global-search-group-heading"><h3>Schedules and alerts</h3><span>${(loadMatches.length + routineMatches.length + alertMatches.length).toLocaleString()} match${loadMatches.length + routineMatches.length + alertMatches.length === 1 ? "" : "es"}</span></div><div class="global-search-tiles">${related.join("")}</div></section>`);
        }

        const resultCount = matchingGroups.length + inventoryMatches.length + clientMatches.length + loadMatches.length + routineMatches.length + alertMatches.length;
        container.classList.remove("hidden");
        container.innerHTML = resultCount
            ? `<div class="global-search-summary"><strong>${resultCount.toLocaleString()} related result${resultCount === 1 ? "" : "s"}</strong><span>Search: “${escapeHtml(String(input.value).trim())}”</span></div>${sections.join("")}`
            : `<div class="global-search-empty"><strong>No matches found for “${escapeHtml(String(input.value).trim())}”.</strong><span>Try a vehicle registration, client, asset type, comment keyword or quantity.</span></div>`;
    }


    function resetAuditPaging() {
        auditLimits.clear();
        auditOlderLimits.clear();
    }

    // The operations day is always UK-local, even when someone opens the
    // dashboard abroad. Unzoned sheet dates retain their written calendar day.
    function auditUkDayKey(value) {
        if (value == null || value === "") return "";
        const raw = String(value).trim();
        const localSheetTimestamp = raw.match(/^(\d{4}-\d{2}-\d{2})(?:[ T].*)?$/);
        const ukSheetTimestamp = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:[ T].*)?$/);
        if (ukSheetTimestamp) return `${ukSheetTimestamp[3]}-${ukSheetTimestamp[2].padStart(2, "0")}-${ukSheetTimestamp[1].padStart(2, "0")}`;
        const serial = Number(raw);
        if (raw && Number.isFinite(serial) && serial > 20000 && serial < 80000) {
            return new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000).toISOString().slice(0, 10);
        }
        if (typeof value === "string" && localSheetTimestamp &&
            !/(?:Z|[+-]\d{2}:?\d{2})$/i.test(raw)) return localSheetTimestamp[1];
        const parsed = value instanceof Date ? value : new Date(value);
        if (Number.isNaN(parsed.getTime())) return localSheetTimestamp?.[1] || "";
        const parts = Object.fromEntries(AUDIT_UK_FORMATTER.formatToParts(parsed)
            .filter(part => ["day", "month", "year"].includes(part.type))
            .map(part => [part.type, part.value]));
        return `${parts.year}-${parts.month}-${parts.day}`;
    }

    function auditSubtractDays(isoDay, days) {
        const date = new Date(`${isoDay}T12:00:00Z`);
        date.setUTCDate(date.getUTCDate() - days);
        return date.toISOString().slice(0, 10);
    }

    function auditDateWindow() {
        const today = auditUkDayKey(new Date());
        const mode = auditDateFilter.mode;
        if (mode === "today") return { from: today, to: today, mode, label: "Today" };
        if (mode === "yesterday") {
            const yesterday = auditSubtractDays(today, 1);
            return { from: yesterday, to: yesterday, mode, label: "Yesterday" };
        }
        if (mode === "7d" || mode === "30d") {
            return { from: auditSubtractDays(today, mode === "7d" ? 6 : 29), to: today,
                mode, label: mode === "7d" ? "Last 7 days" : "Last 30 days" };
        }
        if (mode === "all") return { from: "", to: "", mode, label: "All dates" };
        if (mode === "day") {
            const day = auditDateFilter.day || today;
            return { from: day, to: day, mode, label: `Selected day · ${day}` };
        }
        const from = auditDateFilter.from;
        const to = auditDateFilter.to;
        return { from, to, mode: "custom", label: `${from || "Start"} to ${to || "End"}`,
            error: (!from || !to) ? "Select both the start and end dates." :
                from > to ? "The start date must not be after the end date." : "" };
    }

    function syncAuditDateControls(window = auditDateWindow()) {
        const dayInput = $("dashboard-filter-date-audit");
        if (dayInput) dayInput.value = auditDateFilter.mode === "day" ? auditDateFilter.day : "";
        const range = $("audit-range-fields");
        range?.classList.toggle("hidden", auditDateFilter.mode !== "custom");
        ["from", "to"].forEach(boundary => {
            const input = $("audit-range-" + boundary);
            if (input) input.value = auditDateFilter[boundary];
        });
        document.querySelectorAll("button[data-audit-period]").forEach(button => {
            const active = button.dataset.auditPeriod === auditDateFilter.mode;
            button.classList.toggle("is-active", active);
            button.setAttribute("aria-pressed", String(active));
        });
        const caption = $("audit-period-caption");
        if (caption) caption.textContent = window.mode === "today"
            ? "Today only · use See previous days movements in a category to view its history"
            : `Showing ${window.label.toLowerCase()}`;
        const error = $("audit-range-error");
        if (error) {
            error.textContent = window.error || "";
            error.classList.toggle("hidden", !window.error);
        }
    }

    function renderAudit() {
        const filters = getDashboardFilters();
        const window = auditDateWindow();
        syncAuditDateControls(window);
        // Client and asset remain linked to the dashboard; the audit day/range
        // never modifies or inherits the dashboard-only date filter.
        const matchingRows = getDashboardTransactions().filter(item =>
            (!filters.client || String(item.client || "").trim() === filters.client) &&
            (!filters.asset || String(item.asset || "").trim() === filters.asset)
        );
        const inWindow = item => {
            const key = auditUkDayKey(item.timestamp);
            return !window.error && Boolean(key) &&
                (!window.from || key >= window.from) && (!window.to || key <= window.to);
        };
        const selectedRows = matchingRows.filter(inWindow);
        const regular = matchingRows.filter(item => !FM_DASHBOARD_FEATURES.isAdjustment(item));
        const adjustments = matchingRows.filter(item => FM_DASHBOARD_FEATURES.isAdjustment(item));
        const normalGroups = FM_TRANSACTION_DISPLAY?.groupTransactions?.(regular) || [];
        const adjustmentGroups = FM_TRANSACTION_DISPLAY?.groupTransactions?.(adjustments) || [];
        const types = state.clientTypes || { categories: [], lookup: Object.create(null) };
        const categories = types.categories.length ? types.categories.slice() : ["All clients"];
        const names = new Map(categories.map(name => [name, []]));
        const unmatched = [];
        normalGroups.forEach(group => {
            const assigned = types.lookup[String(group.client || "").trim().replace(/\s+/g, " ").toLocaleLowerCase()];
            if (!types.categories.length) names.get("All clients").push(group);
            else if (names.has(assigned)) names.get(assigned).push(group);
            else unmatched.push(group);
        });
        if (unmatched.length) names.set("Uncategorised", unmatched);
        state.auditVisibleRows = [];
        const count = $("audit-count");
        const visibleCount = [...names.values()].flat().filter(group => inWindow(group)).length +
            adjustmentGroups.filter(group => inWindow(group)).length;
        if (count) count.textContent = visibleCount.toLocaleString("en-GB");

        const renderCards = groups => groups.map(group => {
            const index = state.auditVisibleRows.push(group) - 1;
            return renderTransactionCard(group, index);
        }).join("");

        const renderPanel = (key, title, groups, first) => {
            const open = auditOpen.has(key) ? auditOpen.get(key) : first;
            const limit = auditLimits.get(key) || AUDIT_PAGE_SIZE;
            const current = groups.filter(group => inWindow(group));
            const shown = current.slice(0, limit);
            const todayMode = window.mode === "today" && !window.error;
            const past = todayMode ? groups.filter(group => {
                const day = auditUkDayKey(group.timestamp);
                return day && day < window.from;
            }) : [];
            const olderLimit = auditOlderLimits.get(key) || 0;
            const shownPast = past.slice(0, olderLimit);
            const currentLabel = todayMode ? "today" : "in period";
            const countText = `${current.length.toLocaleString("en-GB")} ${currentLabel}`;
            const empty = todayMode ? "No movements today" : window.error ? "Invalid date range" : "No movements in this period";
            const cards = shown.length
                ? `<div class="transaction-card-list">${renderCards(shown)}</div>`
                : `<div class="audit-empty-period"><strong>${empty}</strong><span>${todayMode
                    ? past.length ? "There are movements from previous days." : "No previous movements are recorded for this category."
                    : "Choose another day or date range to see movements."}</span></div>`;
            const older = shownPast.length
                ? `<div class="audit-older-heading"><strong>Previous days</strong><span>${shownPast.length.toLocaleString("en-GB")} shown</span></div><div class="transaction-card-list">${renderCards(shownPast)}</div>` : "";
            const moreCurrent = current.length > limit
                ? `<button type="button" class="secondary audit-more" data-audit-more="${escapeAttr(key)}">Show more movements (${(current.length - limit).toLocaleString("en-GB")} remaining)</button>` : "";
            // Older entries stay hidden until requested, with pagination to
            // prevent hundreds of old cards from overwhelming mobile devices.
            const morePast = todayMode && past.length && current.length <= limit
                ? `<button type="button" class="secondary audit-more audit-older-button" data-audit-older="${escapeAttr(key)}">${olderLimit
                    ? `Show more previous movements (${(past.length - shownPast.length).toLocaleString("en-GB")} remaining)`
                    : current.length ? `Show previous days movements (${past.length.toLocaleString("en-GB")})`
                        : `See previous days movements (${past.length.toLocaleString("en-GB")})`}</button>` : "";
            const pastButton = shownPast.length >= past.length ? "" : morePast;
            return `<details class="audit-category" data-audit-category="${escapeAttr(key)}" ${open ? "open" : ""}>
                <summary class="audit-category-heading"><span class="audit-category-name">${escapeHtml(title)}</span><span class="audit-category-count">${countText}</span><span class="audit-category-chevron" aria-hidden="true"></span></summary>
                <div class="audit-category-body">${cards}${moreCurrent}${older}${pastButton}</div>
            </details>`;
        };

        const body = $("transactions-body");
        if (body) body.innerHTML = [...names].map(([name, groups], index) => renderPanel(name, name, groups, index === 0)).join("");
        const adjustmentBody = $("audit-adjustments-body");
        if (adjustmentBody) adjustmentBody.innerHTML = renderPanel("__adjustments__", "Normalisations & adjustments", adjustmentGroups, false);

        const selectedRegular = selectedRows.filter(item => !FM_DASHBOARD_FEATURES.isAdjustment(item));
        renderTransactionSummary($("transaction-summary"), selectedRegular, `${window.label} movement totals`);
    }

    function renderTransactionCard(group, index, actionAttr = "data-info-index") {
        const sections = FM_TRANSACTION_DISPLAY.movementSections(group);
        const movementNames = sections.map(section => section.movement);
        const cardClass = movementNames.length === 1
            ? `movement-card-${movementNames[0].toLowerCase()}`
            : movementNames.length > 1 ? "movement-card-mixed" : "movement-card-other";
        const attachment = group.image || group.comment || group.comments?.length ? "Attachment / note" : "";
        const lines = sections.map(section => {
            const label = FM_TRANSACTION_DISPLAY.movementLabel(section.movement);
            const cls = section.movement === "RECEIVED" ? "received" : section.movement === "SENT" ? "sent" : section.movement === "DISCARD" ? "discard" : "other";
            const assets = FM_TRANSACTION_DISPLAY.aggregateAssets(section.items);
            return `<section class="transaction-movement-section">
                <div class="transaction-movement-heading"><span class="transaction-movement-title ${cls}">${escapeHtml(label)}</span><span class="muted">${assets.reduce((sum, item) => sum + item.quantity, 0).toLocaleString("en-GB")} total</span></div>
                <div class="transaction-asset-list">${assets.map(item => `<div class="transaction-asset-pill"><span>${escapeHtml(item.asset)}</span><strong>${formatNumber(item.quantity)}</strong></div>`).join("")}</div>
            </section>`;
        }).join("");
        return `<article class="transaction-card ${cardClass}">
            <div class="transaction-card-head">
                <div class="transaction-card-client"><img class="transaction-truck-icon" src="truck-icon.png" alt=""><div class="transaction-card-client-copy"><div class="eyebrow">${movementNames.includes("NORMALISATION") ? "CLIENT LEDGER ADJUSTMENT" : String(group.client).trim().toLowerCase() === "warehouse" ? "ASSET OPERATION" : "CLIENT MOVEMENT"}</div><h3>${escapeHtml(group.client)}</h3></div></div>
                <time class="transaction-card-time">${escapeHtml(formatTimestamp(group.timestamp))}</time>
            </div>
            <div class="transaction-card-meta"><span class="transaction-card-line-count">${group.totalLines} ledger line${group.totalLines === 1 ? "" : "s"}</span>${group.broker ? `<span class="broker-transaction-tag">Broker · ${escapeHtml(group.broker)}</span>` : ""}${group.vehicleReg ? `<span class="transaction-vehicle-reg"><span>Vehicle reg</span><strong>${escapeHtml(group.vehicleReg)}</strong></span>` : ""}${attachment ? `<span class="transaction-attachment-mark">${attachment}</span>` : ""}</div>
            ${lines}
            <div class="transaction-card-footer"><span class="muted">${group.totalQuantity.toLocaleString("en-GB")} total quantity</span><button type="button" class="secondary" ${actionAttr}="${index}">View details</button></div>
        </article>`;
    }

    function driveFileIdFromLink(url) {
        const text = String(url || "");
        const m = text.match(/\/d\/([a-zA-Z0-9_-]+)/) || text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
        return m ? m[1] : "";
    }

    function prepareInfoModal(group) {
        const summary = $("info-summary");
        if (!summary) return;
        const sections = FM_TRANSACTION_DISPLAY.movementSections(group);
        const comments = Array.isArray(group.comments) && group.comments.length ? group.comments : (group.comment ? [group.comment] : []);
        summary.innerHTML = `<div class="transaction-detail-note"><strong>${escapeHtml(group.client)}</strong>${group.broker ? `<span> · Broker: ${escapeHtml(group.broker)}</span>` : ""}<span> · ${escapeHtml(formatTimestamp(group.timestamp))}</span>${group.vehicleReg ? `<span class="transaction-vehicle-reg"><span>Vehicle reg</span><strong>${escapeHtml(group.vehicleReg)}</strong></span>` : ""}</div>${sections.map(section => {
            const label = FM_TRANSACTION_DISPLAY.movementLabel(section.movement);
            const cls = section.movement === "RECEIVED" ? "received" : section.movement === "SENT" ? "sent" : section.movement === "DISCARD" ? "discard" : "other";
            const assets = FM_TRANSACTION_DISPLAY.aggregateAssets(section.items);
            return `<section class="transaction-detail-section"><div class="transaction-detail-section-head"><span class="transaction-movement-title ${cls}">${escapeHtml(label)}</span><strong>${assets.reduce((sum,item) => sum + item.quantity,0).toLocaleString("en-GB")}</strong></div><div class="transaction-detail-grid">${assets.map(item => `<span>${escapeHtml(item.asset)}</span><strong>${formatNumber(item.quantity)}</strong>`).join("")}</div></section>`;
        }).join("")}${comments.length ? `<div class="transaction-detail-note"><strong>${comments.length === 1 ? "Comment" : "Comments"}</strong><br>${comments.map(comment => escapeHtml(comment)).join("<br>")}</div>` : ""}`;

        const photoWrap = $("info-photo-wrap"), commentEl = $("info-comment"), emptyEl = $("info-empty"), image = $("info-photo"), photoStatus = $("info-photo-status");
        FM_MEDIA?.revokeObjectUrl?.(image);
        image?.removeAttribute("src");
        image?.classList.remove("is-loading");
        commentEl?.classList.add("hidden");
        if (commentEl) commentEl.textContent = "";
        if (group.image) {
            photoWrap?.classList.remove("hidden");
            photoStatus?.classList.remove("hidden");
            if (photoStatus) photoStatus.textContent = "Loading photo…";
            image?.classList.add("is-loading");
            $("info-photo-link").href = group.image;
        } else {
            photoWrap?.classList.add("hidden");
            photoStatus?.classList.add("hidden");
            $("info-photo-link")?.removeAttribute("href");
        }
        if (emptyEl) emptyEl.classList.toggle("hidden", Boolean(group.image || comments.length));
    }

    async function loadInfoPhoto(group) {
        if (!group?.image) return;
        const image = $("info-photo"), photoStatus = $("info-photo-status");
        try {
            const loaded = await FM_MEDIA.loadDriveImage(image, group.image, state.accessToken);
            image?.classList.remove("is-loading");
            photoStatus?.classList.toggle("hidden", loaded);
            if (!loaded && photoStatus) photoStatus.textContent = "Preview unavailable here. Use Open full size in Drive.";
        } catch (error) {
            console.warn("Movement photo preview failed:", error);
            image?.classList.remove("is-loading");
            if (photoStatus) photoStatus.textContent = "Preview unavailable here. Use Open full size in Drive.";
        }
    }

    function openInfoModal(index) { const group = state.auditVisibleRows[index]; if (!group) return; prepareInfoModal(group); $("info-modal").classList.remove("hidden"); requestAnimationFrame(() => loadInfoPhoto(group)); }
    function openClientTransactionInfo(index) { const group = state.clientDetailsRows[index]; if (!group) return; prepareInfoModal(group); $("info-modal").classList.remove("hidden"); requestAnimationFrame(() => loadInfoPhoto(group)); }
    function closeInfoModal() { FM_MEDIA?.revokeObjectUrl?.($("info-photo")); $("info-modal")?.classList.add("hidden"); }

    /* =====================================================
       DASHBOARD
       ===================================================== */


    function renderDashboard() {

        /*
         * NEW:
         * Warehouse asset cards
         */

        renderWarehouseAssets();

        // The filtered load renderer also applies optional Shein estimates.
        renderDashboardViews();


        /* Batch movement form */
        batchForm?.refresh?.();
        handleMovementChange();
        updateMovementPreview();


        $("dashboard-date")
            .textContent =
            new Date().toLocaleDateString(
                "en-GB",
                {

                    weekday:
                        "long",

                    day:
                        "numeric",

                    month:
                        "long",

                    year:
                        "numeric"

                }
            );


        renderInventoryManager();

    }


    /*
     * =====================================================
     * NEW WAREHOUSE ASSET CARDS
     * =====================================================
     *
     * Current number:
     *     state.inventory balance
     *
     * Movement indicator:
     *
     * RECEIVED = +
     * SENT     = -
     * DISCARD  = -
     *
     * Only today's transactions are included.
     */


    function getWarehouseAssetMovements() {

        const movements = {};


        getDashboardTransactions()

            .filter(
                item =>
                    isTodayTransaction(
                        item.timestamp
                    )
            )

            .forEach(
                item => {

                    const asset =
                        String(
                            item.asset ?? ""
                        ).trim();


                    if (!asset) {
                        return;
                    }


                    const movement =
                        String(
                            item.movement ?? ""
                        )
                        .trim()
                        .toUpperCase();


                    const quantity =
                        Number(
                            item.quantity
                        ) || 0;


                    if (
                        !movements[
                            asset.toLowerCase()
                        ]
                    ) {

                        movements[
                            asset.toLowerCase()
                        ] = {

                            asset,

                            net:
                                0,

                            received:
                                0,

                            sent:
                                0,

                            discarded:
                                0

                        };

                    }


                    const entry =
                        movements[
                            asset.toLowerCase()
                        ];


                    if (["RECEIVED", "COMBINE_IN"].includes(movement)) {

                        entry.net +=
                            quantity;

                        entry.received +=
                            quantity;

                    }


                    if (["SENT", "COMBINE_OUT"].includes(movement)) {

                        entry.net -=
                            quantity;

                        entry.sent +=
                            quantity;

                    }


                    if (
                        movement ===
                        "DISCARD"
                    ) {

                        entry.net -=
                            quantity;

                        entry.discarded +=
                            quantity;

                    }

                }
            );


        return movements;

    }


    function renderWarehouseAssets() {
        const container = $("warehouse-assets");
        const filters = getDashboardFilters();
        const assets = getDashboardInventory().filter(item => !filters.asset || String(item.asset).trim() === filters.asset);

        if (!assets.length) {
            container.innerHTML = `<div class="warehouse-empty">No assets match the current filter.</div>`;
            return;
        }

        const activeDate = filters.date || transactionDateKey(new Date().toISOString());

        container.innerHTML = assets.map(item => {
            const assetName = String(item.asset).trim();
            const isHistoricalSnapshot = state.dashboardSource?.kind === "historical";
            const hasExplicitDay = Boolean(filters.date);
            const movementTotal = getDashboardTransactions()
                .filter(t =>
                    String(t.asset || "").trim().toLowerCase() === assetName.toLowerCase() &&
                    (!activeDate || transactionDateKey(t.timestamp) === activeDate) &&
                    (!filters.client || String(t.client || "").trim() === filters.client)
                )
                .reduce((total, t) => {
                    const movement = String(t.movement || "").trim().toUpperCase();
                    const qty = Number(t.quantity) || 0;
                    return ["RECEIVED", "COMBINE_IN"].includes(movement) ? total + qty :
                        ["SENT", "DISCARD", "COMBINE_OUT"].includes(movement) ? total - qty : total;
                }, 0);

            const indicator = isHistoricalSnapshot && !hasExplicitDay
                ? `<span class="asset-movement neutral">—</span>`
                : movementTotal > 0
                    ? `<span class="asset-movement positive">+${formatNumber(movementTotal)}</span>`
                    : movementTotal < 0
                        ? `<span class="asset-movement negative">${formatNumber(movementTotal)}</span>`
                        : `<span class="asset-movement neutral">0</span>`;

            const verifiedBadge = item.verified
                ? (() => {
                    const verificationDate = formatVerificationDate(item.verificationDate);
                    const tooltip = verificationDate
                        ? `Verified on ${verificationDate}`
                        : "Verified";
                    return `<span class="inventory-verified-badge" title="${escapeAttr(tooltip)}" aria-label="${escapeAttr(tooltip)}"><span class="inventory-verified-check" aria-hidden="true">✓</span></span>`;
                })()
                : "";

            return `<div class="warehouse-asset-card" data-asset-name="${escapeAttr(assetName)}" tabindex="0" role="button" aria-label="View ${escapeAttr(assetName)} history">
                <div class="warehouse-asset-name-row">
                    <div class="warehouse-asset-name">${escapeHtml(assetName)}</div>
                    ${verifiedBadge}
                </div>
                <div class="warehouse-asset-total"><div class="warehouse-asset-count">${formatNumber(item.balance)}</div><span>Total</span></div>
                ${(!isHistoricalSnapshot && (!filters.date || filters.date === todayDateKey()) && getTodayAvailability(assetName) !== null)
                    ? `<div class="warehouse-asset-available"><strong>${formatNumber(Math.min(getTodayAvailability(assetName), item.balance))}</strong><span>available to send</span></div>`
                    : ""}
                <div class="warehouse-asset-movement">
                    ${indicator}
                    <span class="movement-label">${filters.date ? formatGraphDate(activeDate) : (state.dashboardSource?.kind === "historical" ? "snapshot" : "today")}</span>
                </div>
            </div>`;
        }).join("");

        // Bind directly to the freshly-rendered cards as a second layer of protection
        // against delegated-event issues after dashboard re-renders.
        container.querySelectorAll("[data-asset-name]").forEach(card => {
            const open = () => openWarehouseAssetDetails(card.dataset.assetName);
            card.addEventListener("click", open);
            card.addEventListener("keydown", event => {
                if (event.key !== "Enter" && event.key !== " " ) return;
                event.preventDefault();
                open();
            });
        });
    }


    function getWarehouseAssetTransactions(asset) {
        const filters = getDashboardFilters();
        const wantedAsset = String(asset || "").trim().toLowerCase();
        return getDashboardTransactions().filter(item => {
            if (String(item.asset || "").trim().toLowerCase() !== wantedAsset) return false;
            if (filters.client && String(item.client || "").trim() !== filters.client) return false;
            return ["SENT", "RECEIVED", "DISCARD", "COMBINE_OUT", "COMBINE_IN"].includes(String(item.movement || "").trim().toUpperCase()) && transactionDateKey(item.timestamp);
        });
    }

    function openWarehouseAssetDetails(asset) {
        const name = String(asset || "").trim();
        if (!name) return;

        const modal = $("warehouse-asset-modal");
        if (!modal) return;

        const inventory = getDashboardInventory().find(item =>
            String(item.asset || "").trim().toLowerCase() === name.toLowerCase()
        );
        const balance = Number(inventory?.balance) || 0;
        const transactions = getWarehouseAssetTransactions(name);
        currentWarehouseGraph = {asset: name, transactions, balance};
        syncGraphRangeControls("movement");
        syncGraphRangeControls("quantity");
        const filters = getDashboardFilters();

        // Open the modal first. Heavy graph work must never block the modal from appearing.
        modal.classList.remove("hidden");
        $("warehouse-asset-modal-title").textContent = name;
        $("warehouse-asset-modal-subtitle").textContent = filters.client
            ? `Warehouse quantity and movements for ${name} · client filter: ${filters.client}`
            : `Warehouse quantity and movements for ${name} · all clients`;

        const received = transactions.filter(item => String(item.movement).trim().toUpperCase() === "RECEIVED")
            .reduce((sum, item) => sum + (Number(item.quantity) || 0), 0);
        const sent = transactions.filter(item => String(item.movement).trim().toUpperCase() === "SENT")
            .reduce((sum, item) => sum + (Number(item.quantity) || 0), 0);
        const discarded = transactions.filter(item => String(item.movement).trim().toUpperCase() === "DISCARD")
            .reduce((sum, item) => sum + (Number(item.quantity) || 0), 0);

        $("warehouse-asset-history-summary").innerHTML = `
            <div class="warehouse-history-stat"><span>Current quantity</span><strong>${formatNumber(balance)}</strong></div>
            <div class="warehouse-history-stat"><span>Total received</span><strong class="stat-positive">${formatNumber(received)}</strong></div>
            <div class="warehouse-history-stat"><span>Total sent</span><strong>${formatNumber(sent)}</strong></div>
            <div class="warehouse-history-stat"><span>Discarded</span><strong>${formatNumber(discarded)}</strong></div>
        `;

        const movementGraph = $("warehouse-asset-movement-graph");
        const quantityGraph = $("warehouse-asset-quantity-graph");
        movementGraph.innerHTML = `<div class="graph-loading-state">Preparing movement history…</div>`;
        quantityGraph.innerHTML = `<div class="graph-loading-state">Preparing quantity history…</div>`;

        requestAnimationFrame(() => {
            try {
                renderWarehouseAssetMovementGraph(name, transactions);
            } catch (error) {
                console.error("Asset movement graph failed:", error);
                movementGraph.innerHTML = `<div class="client-graph-empty"><strong>Movement graph unavailable</strong><span>The asset details are still available above.</span></div>`;
            }
            try {
                renderWarehouseAssetQuantityGraph(name, transactions, balance);
            } catch (error) {
                console.error("Asset quantity graph failed:", error);
                quantityGraph.innerHTML = `<div class="client-graph-empty"><strong>Quantity graph unavailable</strong><span>The asset details are still available above.</span></div>`;
            }
        });
    }

    function closeWarehouseAssetDetails() {
        $("warehouse-asset-modal")?.classList.add("hidden");
    }

    function makeAssetHistorySeries(transactions) {
        const daily = new Map();
        transactions.forEach(item => {
            const date = transactionDateKey(item.timestamp);
            if (!date) return;
            if (!daily.has(date)) daily.set(date, {sent: 0, received: 0, discarded: 0});
            const entry = daily.get(date);
            const movement = String(item.movement || "").trim().toUpperCase();
            const quantity = Number(item.quantity) || 0;
            if (movement === "SENT" || movement === "COMBINE_OUT") entry.sent += quantity;
            else if (movement === "RECEIVED" || movement === "COMBINE_IN") entry.received += quantity;
            else if (movement === "DISCARD") entry.discarded += quantity;
        });
        return [...daily.entries()].sort((a, b) => a[0].localeCompare(b[0]));
    }

    function renderWarehouseAssetMovementGraph(asset, transactions) {
        const graph = $("warehouse-asset-movement-graph");
        const history = makeAssetHistorySeries(transactions);
        if (!history.length) {
            graph.innerHTML = `<div class="client-graph-empty"><strong>No movement history</strong><span>No sent, received or discard transactions were found for ${escapeHtml(asset)}.</span></div>`;
            return;
        }

        const {points, windowInfo, hasMovements} = graphMovementPoints(history, "movement");
        graphWindowSummary(graph, windowInfo, points, "movements");
        if (!hasMovements) {
            graph.innerHTML = `<div class="client-graph-empty"><strong>No movements in this period</strong><span>Select 30 days, 90 days or All to view older transactions.</span></div>`;
            return;
        }
        const width = graphChartWidth(points.length, windowInfo);
        const height = 320, left = 64, right = 28, top = 44, bottom = 62;
        const chartWidth = width - left - right, chartHeight = height - top - bottom;
        const maxValue = niceGraphMax(Math.max(...points.flatMap(([, value]) => [value.sent, value.received, value.discarded]), 1), 4);
        const xFor = i => points.length === 1 ? left + chartWidth / 2 : left + (i / (points.length - 1)) * chartWidth;
        const yFor = value => top + chartHeight - (value / maxValue) * chartHeight;
        const pathFor = key => points.map(([, value], i) => `${i === 0 ? "M" : "L"} ${xFor(i).toFixed(1)} ${yFor(value[key]).toFixed(1)}`).join(" ");
        const yTicks = Array.from({length: 5}, (_, i) => maxValue * (1 - i / 4));
        const grid = yTicks.map((value, i) => {
            const y = top + chartHeight * i / 4;
            return `<line x1="${left}" y1="${y}" x2="${width-right}" y2="${y}" class="graph-grid"></line><text x="${left-10}" y="${y+4}" text-anchor="end" class="graph-y-label">${formatNumber(Math.round(value))}</text>`;
        }).join("");
        const labels = graphXAxisLabels(points, xFor, height-22, windowInfo);
        const pointMarkup = points.map(([date, value], i) => {
            const x = xFor(i), sentY = yFor(value.sent), receivedY = yFor(value.received), discardY = yFor(value.discarded);
            const label = graphPeriodLabel(date, windowInfo.granularity);
            return `<g>
                <circle cx="${x}" cy="${sentY}" r="4.5" class="graph-point-sent"></circle>
                ${graphPointHit(x, sentY, `${label} · Sent: ${formatNumber(value.sent)}`, "graph-point-hit-sent")}
                <circle cx="${x}" cy="${receivedY}" r="4.5" class="graph-point-received"></circle>
                ${graphPointHit(x, receivedY, `${label} · Received: ${formatNumber(value.received)}`, "graph-point-hit-received")}
                ${value.discarded ? `<circle cx="${x}" cy="${discardY}" r="4.5" class="graph-point-discard"></circle>${graphPointHit(x, discardY, `${label} · Discarded: ${formatNumber(value.discarded)}`, "graph-point-hit-discard")}` : ""}
            </g>`;
        }).join("");

        graph.innerHTML = `<div class="client-graph-legend">
            <span class="graph-legend-item"><span class="graph-legend-dot graph-legend-sent"></span>Sent</span>
            <span class="graph-legend-item"><span class="graph-legend-dot graph-legend-received"></span>Received</span>
            <span class="graph-legend-item"><span class="graph-legend-dot graph-legend-discard"></span>Discarded</span>
        </div><div class="client-graph-scroll"><svg class="client-graph-svg" style="--graph-min-width:${width}px" viewBox="0 0 ${width} ${height}" role="img" aria-label="Sent received and discarded quantities for ${escapeHtml(asset)}">
            ${grid}<line x1="${left}" y1="${top+chartHeight}" x2="${width-right}" y2="${top+chartHeight}" class="graph-axis"></line>
            <path d="${pathFor("sent")}" class="graph-line graph-line-sent" fill="none"></path>
            <path d="${pathFor("received")}" class="graph-line graph-line-received" fill="none"></path>
            ${points.some(([,v]) => v.discarded) ? `<path d="${pathFor("discarded")}" class="graph-line graph-line-discard" fill="none"></path>` : ""}
            ${pointMarkup}${labels}</svg></div>`;
        bindGraphTooltips(graph.querySelector(".client-graph-svg"));
    }

    function renderWarehouseAssetQuantityGraph(asset, transactions, currentBalance) {
        const graph = $("warehouse-asset-quantity-graph");
        const daily = makeAssetHistorySeries(transactions);
        if (!daily.length) {
            graph.innerHTML = `<div class="client-graph-empty"><strong>No quantity history</strong><span>A quantity timeline needs transaction history for ${escapeHtml(asset)}.</span></div>`;
            return;
        }

        // Reconstruct historical closing stock by anchoring the series to the
        // current inventory balance and walking the ledger through time.
        const totalNet = daily.reduce((sum, [, value]) => sum + value.received - value.sent - value.discarded, 0);
        let running = currentBalance - totalNet;
        const points = daily.map(([date, value]) => {
            running += value.received - value.sent - value.discarded;
            return [date, Math.max(0, running)];
        });

        const {points: visiblePoints, windowInfo} = graphQuantityPoints(points, "quantity");
        graphWindowSummary(graph, windowInfo, visiblePoints, "closing stock");
        const width = graphChartWidth(visiblePoints.length, windowInfo);
        const height = 320, left = 64, right = 28, top = 44, bottom = 62;
        const chartWidth = width - left - right, chartHeight = height - top - bottom;
        const maxValue = niceGraphMax(Math.max(...visiblePoints.map(([, value]) => value), 1), 4);
        const xFor = i => visiblePoints.length === 1 ? left + chartWidth / 2 : left + (i / (visiblePoints.length - 1)) * chartWidth;
        const yFor = value => top + chartHeight - (value / maxValue) * chartHeight;
        const path = visiblePoints.map(([, value], i) => `${i === 0 ? "M" : "L"} ${xFor(i).toFixed(1)} ${yFor(value).toFixed(1)}`).join(" ");
        const yTicks = Array.from({length: 5}, (_, i) => maxValue * (1 - i / 4));
        const grid = yTicks.map((value, i) => {
            const y = top + chartHeight * i / 4;
            return `<line x1="${left}" y1="${y}" x2="${width-right}" y2="${y}" class="graph-grid"></line><text x="${left-10}" y="${y+4}" text-anchor="end" class="graph-y-label">${formatNumber(Math.round(value))}</text>`;
        }).join("");
        const labels = graphXAxisLabels(visiblePoints, xFor, height-22, windowInfo);
        const pointMarkup = visiblePoints.map(([date, value], i) => {
            const x = xFor(i), y = yFor(value), label = graphPeriodLabel(date, windowInfo.granularity);
            return `<g>
                <circle cx="${x}" cy="${y}" r="5" class="graph-point-quantity"></circle>
                ${graphPointHit(x, y, `${label} · Closing quantity: ${formatNumber(value)}`, "graph-point-hit-quantity")}
            </g>`;
        }).join("");

        graph.innerHTML = `<div class="client-graph-legend"><span class="graph-legend-item"><span class="graph-legend-dot graph-legend-quantity"></span>Closing quantity</span></div>
        <div class="client-graph-scroll"><svg class="client-graph-svg" style="--graph-min-width:${width}px" viewBox="0 0 ${width} ${height}" role="img" aria-label="Calculated quantity over time for ${escapeHtml(asset)}">
            ${grid}<line x1="${left}" y1="${top+chartHeight}" x2="${width-right}" y2="${top+chartHeight}" class="graph-axis"></line>
            <path d="${path}" class="graph-line graph-line-quantity" fill="none"></path>
            ${pointMarkup}${labels}</svg></div>
`;
        bindGraphTooltips(graph.querySelector(".client-graph-svg"));
    }


    function getClientSummary(
        client
    ) {

        const loads =
            state.todayLoads.filter(
                item =>
                    item.client.toLowerCase() ===
                    client.toLowerCase()
            );


        const transactions =
            state.transactions.filter(
                item => {

                    if (
                        String(
                            item.client
                        )
                            .trim()
                            .toLowerCase() !==
                        client.toLowerCase()
                    ) {

                        return false;

                    }


                    return isTodayTransaction(
                        item.timestamp
                    );

                }
            );


        const inboundLoads =
            loads.filter(
                item =>
                    item.direction ===
                    "Inbound"
            );


        const outboundLoads =
            loads.filter(
                item =>
                    item.direction ===
                    "Outbound"
            );


        const received =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "RECEIVED"
            );


        const sent =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "SENT"
            );


        const discarded =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "DISCARD"
            );


        const expectedPallets =
            sumNumericText(
                inboundLoads,
                "pallets"
            );


        const expectedLoose =
            sumNumericText(
                inboundLoads,
                "loose"
            );


        const outboundPallets =
            sumNumericText(
                outboundLoads,
                "pallets"
            );


        const outboundLoose =
            sumNumericText(
                outboundLoads,
                "loose"
            );


        const actualReceived =
            received.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        const actualSent =
            sent.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        const actualDiscarded =
            discarded.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        return {

            client,

            loads,

            inboundLoads,

            outboundLoads,

            transactions,

            received,

            sent,

            discarded,

            expectedPallets,

            expectedLoose,

            outboundPallets,

            outboundLoose,

            actualReceived,

            actualSent,

            actualDiscarded,

            expectedInboundTotal:
                expectedPallets +
                expectedLoose,

            expectedOutboundTotal:
                outboundPallets +
                outboundLoose,

            actualInboundTotal:
                actualReceived,

            actualOutboundTotal:
                actualSent +
                actualDiscarded

        };

    }


    function renderClientCards() {
        const container = $("client-cards");
        const query = String($("client-search")?.value || "").trim().toLowerCase();
        const filters = getDashboardFilters();
        const todayKey = transactionDateKey(new Date().toISOString());

        const clients = getDashboardClients().filter(client => {
            if (query && !client.toLowerCase().includes(query)) return false;
            if (filters.client && client !== filters.client) return false;

            const cardDate = filters.date || todayKey;
            const matchingTransactions = getDashboardTransactions().filter(item =>
                String(item.client || "").trim().toLowerCase() === client.toLowerCase() &&
                transactionDateKey(item.timestamp) === cardDate &&
                (!filters.asset || String(item.asset || "").trim() === filters.asset)
            );

            const canUseLoads = state.dashboardSource?.kind !== "historical" && (!filters.date || filters.date === todayKey);
            const hasLoad = canUseLoads && state.todayLoads.some(load =>
                String(load.client || "").trim().toLowerCase() === client.toLowerCase()
            );

            return !filters.asset || matchingTransactions.length > 0 || hasLoad;
        }).sort((a,b) => a.localeCompare(b));

        container.innerHTML = clients.length ? clients.map(client => {
            const cardDate = filters.date || todayKey;
            const transactions = getDashboardTransactions().filter(item =>
                String(item.client || "").trim().toLowerCase() === client.toLowerCase() &&
                transactionDateKey(item.timestamp) === cardDate &&
                (!filters.asset || String(item.asset || "").trim() === filters.asset)
            );

            const groupedMovementCount = FM_TRANSACTION_DISPLAY?.groupTransactions?.(transactions)?.length || 0;
            const received = transactions.filter(t => String(t.movement || "").toUpperCase() === "RECEIVED")
                .reduce((s,t) => s + (Number(t.quantity) || 0), 0);
            const sent = transactions.filter(t => ["SENT","DISCARD"].includes(String(t.movement || "").toUpperCase()))
                .reduce((s,t) => s + (Number(t.quantity) || 0), 0);
            const assets = [...new Set(transactions.map(t => String(t.asset || "").trim()).filter(Boolean))];

            const canUseLoads = !filters.date || filters.date === todayKey;
            const loads = canUseLoads
                ? state.todayLoads.filter(load => String(load.client || "").trim().toLowerCase() === client.toLowerCase())
                : [];
            const expected = loads.reduce((s, load) => s + sumNumericText([load], "pallets") + sumNumericText([load], "loose"), 0);

            const directionClass = getClientCardDirectionClass({
                inboundLoads: loads.filter(l => l.direction === "Inbound"),
                outboundLoads: loads.filter(l => l.direction === "Outbound")
            });

            return `<article class="client-card client-ledger-card ${directionClass}"
                data-client="${escapeAttr(client)}" tabindex="0" role="button">
                <div class="client-card-top">
                    <div>
                        <div class="eyebrow">CLIENT</div>
                        <h3>${escapeHtml(client)}</h3>
                    </div>
                    <span class="today-pill">${groupedMovementCount} movement${groupedMovementCount === 1 ? "" : "s"}</span>
                </div>
                <div class="client-card-stats">
                    <div><span>Received</span><strong>${formatNumber(received)}</strong></div>
                    <div><span>Sent / discard</span><strong>${formatNumber(sent)}</strong></div>
                    <div><span>Expected today</span><strong>${canUseLoads ? formatNumber(expected) : "—"}</strong></div>
                </div>
                <div class="client-card-assets">
                    ${assets.length
                        ? assets.map(a => `<span class="asset-chip">${escapeHtml(a)}</span>`).join("")
                        : `<span class="muted">No matching asset movements</span>`}
                </div>
                <div class="client-card-footer client-ledger-footer">
                    <span>Asset inventory &amp; transactions</span>
                    <button type="button" class="secondary small-button client-show-more"
                        data-client="${escapeAttr(client)}">Show more</button>
                </div>
            </article>`;
        }).join("") : `
            <div class="client-cards-empty">
                <strong>No clients match the current filters</strong>
                <span>Try clearing a filter or changing the client search.</span>
            </div>`;
    }

    function getClientCardDirectionClass(
        summary
    ) {

        if (
            summary.inboundLoads.length &&
            summary.outboundLoads.length
        ) {

            return "direction-mixed";

        }


        if (
            summary.outboundLoads.length
        ) {

            return "direction-outbound";

        }


        if (
            summary.inboundLoads.length
        ) {

            return "direction-inbound";

        }


        return "direction-neutral";

    }


    function sumNumericText(
        items,
        key
    ) {

        return items.reduce(
            (
                sum,
                item
            ) => {

                const value =
                    String(
                        item[key] ?? ""
                    )
                        .replace(
                            /,/g,
                            ""
                        )
                        .trim();


                const parsed =
                    Number(value);


                return Number.isFinite(
                    parsed
                )
                    ? sum + parsed
                    : sum;

            },
            0
        );

    }


    function renderLoadTable(body, items) {
        body.innerHTML = items.length
            ? items.map(item => `
                <tr>
                    <td><strong>${escapeHtml(item.client || "Unknown")}</strong></td>
                    <td>
                        <strong>${item.experimentalEstimate ? "Estimated" : escapeHtml(item.plannedArrival || "Routine")}</strong>
                        ${item.experimentalEstimate ? `<span class="subtext">Since ${escapeHtml(formatTimestamp(item.lastSent))}</span>` : item.loadType ? `<span class="subtext">${escapeHtml(item.loadType)}</span>` : ""}
                    </td>
                    <td><strong>${escapeHtml(item.asset || (/pallet/i.test(item.asset || "") ? "Pallets" : "Assets"))}</strong></td>
                    <td class="num">${formatNumber(item.quantity || 0)}</td>
                    <td>${escapeHtml(item.destination || "—")}</td>
                </tr>`).join("")
            : emptyRow(5, "No routine movements scheduled for today.");
    }


    async function recordMovement(event) {
        event.preventDefault();
        if (movementSubmitBusy) return;

        const entries = batchForm?.read?.() || [];
        const validationError = batchForm?.validate?.();
        const timestampInput = $("timestamp")?.value || "";
        const user = state.idTokenPayload?.email || readSavedSession()?.email || state.idTokenPayload?.name || "Google user";
        const comment = $("movement-comment")?.value.trim() || "";
        const photoFile = $("movement-photo")?.files?.[0] || window._fmCameraFile || null;

        if (validationError) { setMovementStatus(validationError, true); return; }
        if (!timestampInput) { setMovementStatus("Choose a time for this movement batch.", true); return; }

        // Quick local validation for instant feedback. The live Sheet is re-read
        // again immediately before the atomic write, so a stale page snapshot
        // can never be used to overwrite another user's balance.
        const balances = new Map(state.inventory.map(item => [String(item.asset || "").toLowerCase(), Number(item.balance) || 0]));
        const prepared = [];
        for (let entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
            const entry = entries[entryIndex];
            const client = entry.movement === "DISCARD" ? (CONFIG.SELF_CLIENT_NAME || "HSC London (Self)") : entry.client;
            const seen = new Set();
            for (const assetEntry of entry.items) {
                const key = String(assetEntry.asset || "").toLowerCase();
                if (seen.has(key)) { setMovementStatus(`Client ${entryIndex + 1} has ${assetEntry.asset} more than once.`, true); return; }
                seen.add(key);
                const item = state.inventory.find(x => String(x.asset || "").toLowerCase() === key);
                if (!item) { setMovementStatus(`Asset "${assetEntry.asset}" is not present in Inventory.`, true); return; }
                const current = balances.get(key) || 0;
                const next = entry.movement === "RECEIVED" ? current + assetEntry.quantity : current - assetEntry.quantity;
                if (entry.movement !== "RECEIVED" && next < 0) { setMovementStatus(`Cannot remove ${assetEntry.quantity} ${item.asset}. Current balance is ${formatNumber(current)}.`, true); return; }
                balances.set(key, next);
                prepared.push({ client, broker: entry.movement === "DISCARD" ? "" : (entry.broker || ""), movement: entry.movement, asset: item.asset, quantity: assetEntry.quantity });
            }
        }

        movementSubmitBusy = true;
        setMovementSubmitting(true, "Checking Google access...");
        try { await ensureWriteAccess(); }
        catch (authError) {
            movementSubmitBusy = false; setMovementSubmitting(false);
            showReconnectUI("Google access expired. Allow Sheets & Drive access, then submit the batch again.");
            setMovementStatus(authError.message || "Google access could not be refreshed.", true); return;
        }

        try {
            const ledgerId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
            const timestamp = transactionTimestampFromTime(timestampInput);
            let imageLink = "";
            if (photoFile) {
                setMovementSubmitting(true, "Uploading photo...");
                setMovementStatus("Uploading photo...");
                const folderId = await getOrCreateDailyFolder(transactionDateKey(timestamp));
                const uploadFile = await FM_MEDIA.optimizeImageForUpload(photoFile);
                const uploaded = await uploadImageToDrive(uploadFile, folderId, buildPhotoFilename({entries, prepared, timestamp, photoFile}));
                imageLink = uploaded.webViewLink || `https://drive.google.com/file/d/${uploaded.id}/view`;
            }

            // Read the live balance once before the write. The writer then records the
            // resulting warehouse balance directly on each transaction row and
            // updates Asset Inventory in the same Sheets batchUpdate request.
            setMovementSubmitting(true, "Checking live inventory...");
            const latestInventory = parseInventory(await getValues(ledgerId, CONFIG.INVENTORY_SHEET_NAME));
            const latestByKey = new Map(latestInventory.map(item => [String(item.asset || "").trim().toLowerCase(), item]).filter(([key]) => key));
            const liveBalances = new Map(latestInventory.map(item => [String(item.asset || "").trim().toLowerCase(), Number(item.balance) || 0]).filter(([key]) => key));
            const livePrepared = [];

            for (const movement of prepared) {
                const key = String(movement.asset || "").trim().toLowerCase();
                const item = latestByKey.get(key);
                if (!item) throw new Error(`Asset "${movement.asset}" is no longer present in Inventory. Refresh the page and try again.`);
                const current = Number(liveBalances.get(key) || 0);
                const next = movement.movement === "RECEIVED" ? current + movement.quantity : current - movement.quantity;
                if (movement.movement !== "RECEIVED" && next < 0) {
                    throw new Error(`The live balance for ${item.asset} is now ${formatNumber(current)}. Another user may have moved some stock. Refresh the page and submit again.`);
                }
                liveBalances.set(key, next);
                livePrepared.push({ ...movement, asset: item.asset, item });
            }

            if (!window.FM_SHEET_WRITER?.appendMovementBatch) {
                throw new Error("The sheet writer is unavailable. Refresh the page and try again.");
            }

            setMovementSubmitting(true, "Saving movement...");
            const result = await FM_SHEET_WRITER.appendMovementBatch({
                spreadsheetId: ledgerId,
                accessToken: state.accessToken,
                movements: livePrepared.map(item => ({
                    timestamp,
                    client: item.client,
                    broker: item.broker || "",
                    movement: item.movement,
                    asset: item.asset,
                    quantity: item.quantity,
                    user,
                    comment,
                    image: imageLink
                })),
                user,
                timestamp
            });
            if (!result?.movements?.length) throw new Error("Google Sheets did not confirm the movement write.");

            resetMovementForm();
            setMovementStatus(`${entries.length} client${entries.length === 1 ? "" : "s"} · ${livePrepared.length} movement${livePrepared.length === 1 ? "" : "s"} recorded successfully.`);
            await loadDashboard();
        } catch (error) {
            console.error(error);
            setMovementStatus(error.message || "Unable to record movement batch.", true);
        } finally {
            movementSubmitBusy = false;
            setMovementSubmitting(false);
        }
    }

    let movementSubmitLoadingTimer = null;
    function setMovementSubmitting(busy, detail = "") {
        const overlay = $("movement-submit-loading");
        if (movementSubmitLoadingTimer) {
            clearTimeout(movementSubmitLoadingTimer);
            movementSubmitLoadingTimer = null;
        }
        if (overlay && busy) {
            movementSubmitLoadingTimer = setTimeout(() => {
                if (!movementSubmitBusy) return;
                overlay.classList.remove("hidden");
                overlay.setAttribute("aria-hidden", "false");
            }, 250);
        } else if (overlay) {
            overlay.classList.add("hidden");
            overlay.setAttribute("aria-hidden", "true");
        }
        if ($("movement-submit-loading-detail") && detail) $("movement-submit-loading-detail").textContent = detail;
        const submit = $("movement-form")?.querySelector("button[type=submit]");
        if (submit) {
            if (!submit.dataset.originalText) submit.dataset.originalText = submit.textContent;
            submit.disabled = busy;
            submit.textContent = busy ? "Saving..." : submit.dataset.originalText;
        }
        batchForm?.setAllDisabled?.(busy);
    }

    function handleDashboardPhotoChange() {
        const file = $("movement-photo")?.files?.[0];
        const wrap = $("photo-preview-wrap");
        const clear = $("clear-photo");
        if (!file) {
            wrap?.classList.add("hidden");
            clear?.classList.add("hidden");
            return;
        }
        $("photo-preview").src = URL.createObjectURL(file);
        wrap?.classList.remove("hidden");
        clear?.classList.remove("hidden");
    }

    function clearDashboardPhoto() {
        const input = $("movement-photo");
        if (input) input.value = "";
        if ($("movement-photo-camera")) $("movement-photo-camera").value = "";
        $("photo-preview-wrap")?.classList.add("hidden");
        $("photo-preview")?.removeAttribute("src");
        $("clear-photo")?.classList.add("hidden");
    }

    async function getOrCreateDailyFolder(dateStr) {
        if (!window._fmDriveFolderCache) window._fmDriveFolderCache = {};
        if (window._fmDriveFolderCache[dateStr]) return window._fmDriveFolderCache[dateStr];

        const parent = CONFIG.DRIVE_PHOTOS_PARENT_FOLDER_ID;
        const query = [
            `name = '${escapeDriveQuery(dateStr)}'`,
            `mimeType = 'application/vnd.google-apps.folder'`,
            `'${parent}' in parents`,
            'trashed = false'
        ].join(" and ");

        const found = await fetchJson(
            `${DRIVE_API}?q=${encodeURIComponent(query)}&pageSize=1&fields=files(id,name)`,
            {headers: authHeaders()}
        );

        let id = found.files?.[0]?.id;
        if (!id) {
            const created = await fetchJson(DRIVE_API, {
                method: "POST",
                headers: {...authHeaders(), "Content-Type": "application/json"},
                body: JSON.stringify({name: dateStr, mimeType: "application/vnd.google-apps.folder", parents: [parent]})
            });
            id = created.id;
        }
        window._fmDriveFolderCache[dateStr] = id;
        return id;
    }

    async function uploadImageToDrive(file, folderId, filename) {
        const metadata = {name: filename, parents: [folderId], mimeType: file.type || "image/jpeg"};
        const boundary = "fmdashboard" + Math.random().toString(36).slice(2);
        const delimiter = `--${boundary}\r\n`;
        const closeDelim = `\r\n--${boundary}--`;
        const metaPart = delimiter + "Content-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(metadata) + "\r\n";
        const mediaHeader = delimiter + `Content-Type: ${metadata.mimeType}\r\n\r\n`;
        const body = new Blob([metaPart, mediaHeader, file, closeDelim]);

        return fetchJson(
            `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink`,
            {
                method: "POST",
                headers: {...authHeaders(), "Content-Type": `multipart/related; boundary=${boundary}`},
                body
            }
        );
    }

    function buildPhotoFilename(data = {}) {
        const records = Array.isArray(data.prepared) && data.prepared.length
            ? data.prepared
            : (Array.isArray(data.entries) ? data.entries : []);
        const clientNames = [...new Set(records.map(record => {
            const movement = String(record?.movement || "").toUpperCase();
            return String(movement === "DISCARD" ? (CONFIG.SELF_CLIENT_NAME || "HSC London (Self)") : (record?.client || "")).trim();
        }).filter(Boolean))];
        const fallbackClient = String(data.client || "").trim();
        if (!clientNames.length && fallbackClient && !["batch", "unknown client", "client"].includes(fallbackClient.toLowerCase())) {
            clientNames.push(fallbackClient);
        }
        const clientPart = clientNames.length
            ? (clientNames.length <= 3 ? clientNames.join("-AND-") : `${clientNames[0]}-AND-${clientNames.length - 1}-OTHERS`)
            : "UNKNOWN-CLIENT";
        const safeClient = clientPart.toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 90) || "UNKNOWN-CLIENT";
        const movementNames = [...new Set(records.map(record => String(record?.movement || "").trim().toUpperCase()).filter(Boolean))];
        const suppliedMovement = String(data.movement || "").trim().toUpperCase();
        if (!movementNames.length && suppliedMovement && suppliedMovement !== "MIXED") movementNames.push(suppliedMovement);
        const movementPart = movementNames.length === 1
            ? (movementNames[0] === "DISCARD" ? "DISCARDED" : movementNames[0])
            : "MIXED";
        const suppliedDate = data.timestamp ? new Date(data.timestamp) : null;
        const fileDate = suppliedDate && !Number.isNaN(suppliedDate.getTime()) ? suppliedDate : new Date();
        const stamp = fileDate.toISOString().replace(/[:.]/g, "-");
        const uniqueSuffix = Date.now().toString(36).toUpperCase();
        const extension = data.photoFile?.type === "image/jpeg"
            ? ".jpg"
            : ((data.photoFile?.name?.match(/\.[a-zA-Z0-9]+$/) || [".jpg"])[0]);
        return `${safeClient}-${movementPart}-${stamp}-${uniqueSuffix}${extension}`;
    }


    function movementLabel(value) {
        switch (String(value || '').toUpperCase()) {
            case 'RECEIVED':
                return 'Received';
            case 'SENT':
                return 'Sent';
            case 'DISCARD':
                return 'Discard';
            case 'COMBINE_OUT':
                return 'Combination component used';
            case 'COMBINE_IN':
                return 'Combined asset created';
            case 'NORMALISATION':
                return 'Client ledger normalisation';
            default:
                return String(value || '');
        }
    }


    function updateMovementPreview() {
        batchForm?.updatePreview?.("preview-text");
    }

    function resetMovementForm() {
        $("movement-form")?.reset();
        batchForm?.reset?.();
        setDefaultTimestamp();
        handleMovementChange();
        clearDashboardPhoto();
        window._fmCameraFile = null;
    }


    async function addAssetType() {

        const input =
            $("new-asset-name");


        const balanceInput =
            $("new-asset-balance");


        const assetName =
            input.value.trim();


        const balance =
            Number(
                balanceInput.value
            );


        if (!assetName) {

            setInventoryManageStatus(
                "Enter an asset type name.",
                true
            );

            return;

        }


        if (
            !Number.isInteger(
                balance
            ) ||
            balance < 0
        ) {

            setInventoryManageStatus(
                "Starting balance must be a whole number of zero or more.",
                true
            );

            return;

        }


        const exists =
            state.inventory.some(
                item =>
                    item.asset
                        .toLowerCase() ===
                    assetName.toLowerCase()
            );


        if (exists) {

            setInventoryManageStatus(
                "That asset type already exists.",
                true
            );

            return;

        }


        try {

            setInventoryManageStatus(
                "Adding asset type..."
            );


            if (!window.FM_SHEET_WRITER?.appendAsset) {
                throw new Error("The sheet writer is unavailable. Refresh the page and try again.");
            }

            await FM_SHEET_WRITER.appendAsset({
                spreadsheetId: CONFIG.INVENTORY_LEDGER_SHEET_ID,
                accessToken: state.accessToken,
                assetName,
                startingBalance: balance
            });


            input.value = "";

            balanceInput.value =
                "0";


            setInventoryManageStatus(
                "Asset type added successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    async function deleteAssetType(
        rowNumber
    ) {

        const item =
            state.inventory.find(
                inventoryItem =>
                    inventoryItem.rowNumber ===
                    rowNumber
            );


        if (!item) {
            return;
        }


        if (
            item.balance !== 0
        ) {

            setInventoryManageStatus(
                `Cannot remove ${item.asset} because its current balance is ${item.balance}. Set the balance to zero first.`,
                true
            );

            return;

        }


        const confirmed =
            window.confirm(
                `Remove the asset type "${item.asset}" from inventory?`
            );


        if (!confirmed) {
            return;
        }


        try {

            setInventoryManageStatus(
                "Removing asset type..."
            );


            const metadata =
                await sheetsGet(
                    `/${encodeURIComponent(
                        CONFIG.INVENTORY_LEDGER_SHEET_ID
                    )}`
                );


            const sheet =
                (metadata.sheets || [])
                    .find(
                        sheet =>
                            sheet.properties.title ===
                            CONFIG.INVENTORY_SHEET_NAME
                    );


            if (!sheet) {

                throw new Error(
                    `Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`
                );

            }


            const sheetId =
                sheet.properties.sheetId;


            await sheetsPost(

                `/${encodeURIComponent(
                    CONFIG.INVENTORY_LEDGER_SHEET_ID
                )}:batchUpdate`,

                {

                    requests: [

                        {

                            deleteDimension: {

                                range: {

                                    sheetId,

                                    dimension:
                                        "ROWS",

                                    startIndex:
                                        rowNumber - 1,

                                    endIndex:
                                        rowNumber

                                }

                            }

                        }

                    ]

                }

            );


            setInventoryManageStatus(
                "Asset type removed successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    async function renameAssetType(
        rowNumber
    ) {

        const item =
            state.inventory.find(
                inventoryItem =>
                    inventoryItem.rowNumber ===
                    rowNumber
            );


        if (!item) {
            return;
        }


        const newName =
            window.prompt(
                `Rename "${item.asset}" to:`,
                item.asset
            );


        if (
            newName === null
        ) {
            return;
        }


        const cleanName =
            newName.trim();


        if (!cleanName) {

            setInventoryManageStatus(
                "Asset name cannot be empty.",
                true
            );

            return;

        }


        const duplicate =
            state.inventory.some(
                other =>
                    other.rowNumber !==
                        rowNumber &&

                    other.asset
                        .toLowerCase() ===
                    cleanName.toLowerCase()
            );


        if (duplicate) {

            setInventoryManageStatus(
                "Another asset already has that name.",
                true
            );

            return;

        }


        try {

            setInventoryManageStatus(
                "Renaming asset type..."
            );


            if (!window.FM_SHEET_WRITER?.migrateRename) {
                throw new Error("The sheet writer is unavailable. Refresh the page and try again.");
            }

            await FM_SHEET_WRITER.migrateRename({
                spreadsheetId: CONFIG.INVENTORY_LEDGER_SHEET_ID,
                accessToken: state.accessToken,
                newName: cleanName,
                rowNumber: item.rowNumber,
                assetColumn: item.assetColumn
            });


            setInventoryManageStatus(
                "Asset type renamed successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    function openInventoryManager() {

        renderInventoryManager();


        $("inventory-modal")
            .classList
            .remove("hidden");

    }


    function closeInventoryManager() {

        $("inventory-modal")
            .classList
            .add("hidden");

    }


    function renderInventoryManager() {

        const body =
            $("manage-inventory-body");


        if (!body) {
            return;
        }


        body.innerHTML =

            state.inventory.length

                ? state.inventory.map(
                    item => `

                        <tr>

                            <td>

                                <strong>
                                    ${escapeHtml(
                                        item.asset
                                    )}
                                </strong>

                            </td>

                            <td class="num">${formatNumber(item.balance)}</td>
                            <td>
                                <input class="availability-count-input" type="number" min="0" step="1"
                                    data-availability-asset="${escapeAttr(item.asset)}"
                                    value="${getTodayAvailability(item.asset) === null ? "" : Math.min(getTodayAvailability(item.asset), item.balance)}"
                                    placeholder="—" aria-label="Available ${escapeAttr(item.asset)} to send today">
                            </td>
                            <td class="actions-cell">

                                <button
                                    class="table-action"
                                    data-action="rename"
                                    data-row="${item.rowNumber}"
                                >
                                    Rename
                                </button>


                                <button
                                    class="table-action danger"
                                    data-action="delete"
                                    data-row="${item.rowNumber}"
                                >
                                    Remove
                                </button>

                            </td>

                        </tr>

                    `
                ).join("")

                : emptyRow(4, "No asset types configured.");

    }


    async function saveAvailabilityCounts() {
        if (!state.accessToken) return;
        const inputs = [...document.querySelectorAll("[data-availability-asset]")];
        const entries = [];
        for (const input of inputs) {
            const asset = String(input.dataset.availabilityAsset || "").trim();
            const raw = String(input.value ?? "").trim();
            if (!asset) continue;
            if (!raw) { entries.push({asset, count: null}); continue; }
            const count = Number(raw);
            const item = state.inventory.find(row => row.asset.toLowerCase() === asset.toLowerCase());
            if (!Number.isInteger(count) || count < 0) { setInventoryManageStatus(`Enter a whole number of zero or more for ${asset}, or leave it blank.`, true); input.focus(); return; }
            if (item && count > item.balance) { setInventoryManageStatus(`${asset}: available to send cannot exceed the total stock (${formatNumber(item.balance)}).`, true); input.focus(); return; }
            entries.push({asset, count});
        }
        try {
            setInventoryManageStatus("Saving availability counts…");
            const ledgerId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
            const rows = await getValues(ledgerId, CONFIG.ASSET_AVAILABILITY_SHEET_NAME);
            const existing = parseAvailability(rows);
            const date = todayDateKey();
            const user = state.idTokenPayload?.email || readSavedSession()?.email || "";
            const timestamp = new Date().toISOString();
            const header = rows[0] || ["Date", "Asset", "Available Count", "Updated By", "Updated At"];
            const normalized = header.map(normalizeHeader);
            const dateIdx = Math.max(0, normalized.indexOf("date"));
            const assetIdx = Math.max(0, normalized.indexOf("asset"));
            const byIdx = Math.max(0, normalized.indexOf("updated by"));
            const atIdx = Math.max(0, normalized.indexOf("updated at"));
            const countIdx = Math.max(0, normalized.indexOf("available count"));
            const updates = [];
            const clears = [];
            let nextRow = Math.max(rows.length, 1) + 1;
            for (const entry of entries) {
                const old = existing.filter(row => row.date === date && row.asset.toLowerCase() === entry.asset.toLowerCase()).slice(-1)[0];
                if (entry.count === null) {
                    if (old) clears.push(`'${CONFIG.ASSET_AVAILABILITY_SHEET_NAME}'!A${old.rowNumber}:E${old.rowNumber}`);
                    continue;
                }
                let rowNumber = old?.rowNumber || nextRow++;
                const values = new Array(Math.max(header.length, 5)).fill("");
                values[dateIdx] = date; values[assetIdx] = entry.asset; values[countIdx] = entry.count; values[byIdx] = user; values[atIdx] = timestamp;
                updates.push({range: `'${CONFIG.ASSET_AVAILABILITY_SHEET_NAME}'!A${rowNumber}:${columnLetter(values.length)}${rowNumber}`, majorDimension: "ROWS", values: [values]});
            }
            if (updates.length) await sheetsPost(`/${encodeURIComponent(ledgerId)}/values:batchUpdate`, {valueInputOption: "USER_ENTERED", data: updates});
            for (const range of clears) await sheetsPost(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(range)}:clear`, {});
            const refreshed = await getValues(ledgerId, CONFIG.ASSET_AVAILABILITY_SHEET_NAME);
            state.assetAvailability = parseAvailability(refreshed);
            renderInventoryManager();
            renderWarehouseAssets();
            setInventoryManageStatus("Availability counts saved.");
        } catch (error) {
            console.error(error);
            setInventoryManageStatus(error.message || "Could not save availability counts.", true);
        }
    }

    function handleMovementChange() {
        batchForm?.refresh?.();
        updateMovementPreview();
    }


    function transactionTimestampFromTime(timeValue) {
        const value = String(timeValue || "").trim();
        if (!/^\d{2}:\d{2}$/.test(value)) return new Date().toISOString();
        const [hours, minutes] = value.split(":").map(Number);
        const now = new Date();
        const selected = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hours, minutes, 0, 0);
        const offset = -selected.getTimezoneOffset();
        const sign = offset >= 0 ? "+" : "-";
        const abs = Math.abs(offset);
        const offsetHours = String(Math.floor(abs / 60)).padStart(2, "0");
        const offsetMinutes = String(abs % 60).padStart(2, "0");
        const pad = value => String(value).padStart(2, "0");
        return `${selected.getFullYear()}-${pad(selected.getMonth() + 1)}-${pad(selected.getDate())}T${pad(hours)}:${pad(minutes)}:00${sign}${offsetHours}:${offsetMinutes}`;
    }

    function setDefaultTimestamp() {

        const input =
            $("timestamp");


        if (!input) {
            return;
        }


        const now =
            new Date();


        input.value =

            `${String(
                now.getHours()
            ).padStart(2, "0")}:${String(
                now.getMinutes()
            ).padStart(2, "0")}`;

    }


    function updateAlertNavBadge() {
        const count = state.alerts.filter(alert => alert?.status === "OPEN").length;
        document.querySelectorAll("#alerts-nav-badge").forEach(el => {
            el.textContent = count.toLocaleString();
            el.classList.toggle("hidden", count === 0);
        });
    }

    function updateAlertNavBadges() { updateAlertNavBadge(); }

    function signOut() {
        const saved = readSavedSession();
        if (state.idTokenPayload?.sub) { try { google.accounts.id.revoke(state.idTokenPayload.sub, () => {}); } catch (_) {} }
        if (saved?.sub && saved.sub !== state.idTokenPayload?.sub) { try { google.accounts.id.revoke(saved.sub, () => {}); } catch (_) {} }
        localStorage.removeItem(SESSION_KEY);
        window.FM_AUTH_CACHE?.clear?.();
        state.idTokenPayload = null; state.accessToken = null; state.workbookId = null; state.workbookName = null; state.inventory = []; state.transactions = []; state.dashboardInventory = []; state.dashboardTransactions = []; state.dashboardClients = []; state.historicalLedgerFiles = []; state.dashboardSource = {kind:"current", id:null, name:"Current live ledger"}; state.todayLoads = []; state.routines = []; state.alerts = []; state.clients = []; state.clientTypes = { categories: [], lookup: Object.create(null) };
        $("dashboard").classList.add("hidden"); $("login-card").classList.remove("hidden"); $("google-signin-button").classList.remove("hidden"); $("grant-access").classList.add("hidden"); $("sign-out").classList.add("hidden"); $("user-photo").classList.add("hidden"); $("user-name").textContent = "Not signed in"; $("user-email").textContent = "";
        closeInventoryManager();
    }


    function saveSession() { if (state.idTokenPayload) localStorage.setItem(SESSION_KEY, JSON.stringify({ name: state.idTokenPayload.name || "Google user", email: state.idTokenPayload.email || "", picture: state.idTokenPayload.picture || "", sub: state.idTokenPayload.sub || "" })); }
    function readSavedSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch (_) { return null; } }
    function setUserProfile(profile) { const rawName = String(profile?.given_name || profile?.name || "Google user").trim(); const firstName = (rawName.split(/\s+/)[0] || "there").replace(/[^\p{L}\p{M}'-]/gu, ""); $("user-name").textContent = `Hi, ${firstName || "there"} 👋`; $("user-email").textContent = ""; if (profile?.picture) { $("user-photo").src = profile.picture; $("user-photo").classList.remove("hidden"); } }
    function hideLogin() { $("google-signin-button").classList.add("hidden"); $("grant-access").classList.add("hidden"); $("sign-out").classList.remove("hidden"); $("login-card").classList.add("hidden"); $("dashboard").classList.remove("hidden"); }

    function showReconnectUI(message = "Google access expired. Allow Sheets & Drive access to continue.") {
        window.FM_CONNECTION_UI?.hide();
        $("login-card")?.classList.remove("hidden");
        $("google-signin-button")?.classList.add("hidden");
        $("grant-access")?.classList.remove("hidden");
        const title = $("login-card")?.querySelector("h2");
        if (title) title.textContent = "Reconnect Google Sheets";
        setAuthStatus(message, true);
    }

    function refreshSavedSessionIfNeeded() {
        const saved = readSavedSession();
        if (!saved || !window.google?.accounts?.oauth2) return;
        if (window.FM_AUTH_CACHE?.read?.(saved.email)?.token) return;
        attemptSilentAccess(saved.email);
    }

    window.addEventListener("pageshow", refreshSavedSessionIfNeeded);
    document.addEventListener("visibilitychange", () => {
        if (!document.hidden) refreshSavedSessionIfNeeded();
    });


    function authHeaders() {

        return {

            Authorization:
                `Bearer ${state.accessToken}`

        };

    }


    async function sheetsGet(
        path
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                headers:
                    authHeaders()

            }

        );

    }


    async function sheetsPost(
        path,
        body
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                method:
                    "POST",

                headers: {

                    ...authHeaders(),

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify(
                        body
                    )

            }

        );

    }


    async function sheetsPut(
        path,
        body
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                method:
                    "PUT",

                headers: {

                    ...authHeaders(),

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify(
                        body
                    )

            }

        );

    }


    function isTransientApiStatus(status) { return [408, 429, 500, 502, 503, 504].includes(Number(status)); }

    async function fetchJson(url, options = {}) {
        const method = String(options.method || "GET").toUpperCase();
        const safeRetry = ["GET", "HEAD"].includes(method);
        const maxRetries = safeRetry ? 3 : 0;
        let response;

        for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
            response = await fetch(url, options);

            if (response.status === 401 && !options.__retried) {
                try {
                    await acquireAccessToken("none", state.idTokenPayload?.email || readSavedSession()?.email, {forceRefresh: true});
                    return fetchJson(url, { ...options, __retried: true, headers: { ...(options.headers || {}), ...authHeaders() } });
                } catch (_) {
                    state.accessToken = null;
                    window.FM_AUTH_CACHE?.clear?.();
                    showReconnectUI("Google Sheets access expired. Allow Sheets & Drive access to reconnect.");
                    throw new Error("Google Sheets access expired. Allow Sheets & Drive access to reconnect.");
                }
            }

            if (response.ok) break;
            if (!safeRetry || !isTransientApiStatus(response.status) || attempt >= maxRetries) break;

            const retryAfter = Number(response.headers?.get?.("Retry-After"));
            const delay = Number.isFinite(retryAfter) && retryAfter > 0
                ? Math.min(32000, retryAfter * 1000)
                : Math.min(8000, 1000 * (2 ** attempt)) + Math.floor(Math.random() * 250);
            await new Promise(resolve => setTimeout(resolve, delay));
        }

        const text = await response.text();
        let data = {};
        try { data = text ? JSON.parse(text) : {}; } catch (_) {}
        if (!response.ok) {
            const error = new Error(data?.error?.message || `Request failed (${response.status})`);
            error.status = response.status;
            throw error;
        }
        return data;
    }


    function decodeJwtPayload(
        jwt
    ) {

        const part =
            jwt.split(".")[1];


        const normalized =
            part
                .replace(
                    /-/g,
                    "+"
                )
                .replace(
                    /_/g,
                    "/"
                );


        const padded =
            normalized +
            "=".repeat(

                (
                    4 -
                    normalized.length % 4
                ) % 4

            );


        return JSON.parse(

            decodeURIComponent(

                Array.from(
                    atob(padded)
                )

                .map(
                    character =>
                        `%${character
                            .charCodeAt(0)
                            .toString(16)
                            .padStart(2, "0")}`
                )

                .join("")

            )

        );

    }


    function normalizeHeader(
        value
    ) {

        return String(
            value ?? ""
        )
            .trim()
            .toLowerCase()
            .replace(
                /\s+/g,
                " "
            );

    }


    function findColumn(
        header,
        names
    ) {

        const normalized =
            names.map(
                normalizeHeader
            );


        return header.findIndex(
            value =>
                normalized.includes(
                    value
                )
        );

    }


    function isTruthyVerificationValue(value) {
        if (value === true || value === 1) return true;
        const normalized = String(value ?? "").trim().toLowerCase();
        return ["true", "1", "yes", "y"].includes(normalized);
    }


    function formatVerificationDate(value) {
        if (value === null || value === undefined || String(value).trim() === "") return "";

        const raw = String(value).trim();
        const numeric = Number(raw);
        const date = Number.isFinite(numeric) && numeric > 20000 && numeric < 100000
            ? new Date((numeric - 25569) * 86400 * 1000)
            : new Date(raw);

        if (Number.isNaN(date.getTime())) return raw;

        return date.toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short",
            year: "numeric"
        });
    }


    function numericValue(
        value
    ) {

        const text =
            String(
                value ?? ""
            )
                .replace(
                    /,/g,
                    ""
                )
                .trim();


        if (!text) {
            return 0;
        }


        const number =
            Number(text);


        return Number.isFinite(
            number
        )
            ? number
            : 0;

    }


    function formatNumber(
        value
    ) {

        return Number(
            value || 0
        )
            .toLocaleString(
                "en-GB"
            );

    }


    function quoteSheetName(
        sheetName
    ) {

        return "'" +

            String(
                sheetName
            )
                .replace(
                    /'/g,
                    "''"
                ) +

            "'";

    }


    function columnLetter(
        number
    ) {

        let result =
            "";


        while (
            number > 0
        ) {

            const remainder =
                (
                    number - 1
                ) % 26;


            result =
                String.fromCharCode(
                    65 + remainder
                ) +
                result;


            number =
                Math.floor(
                    (
                        number - 1
                    ) / 26
                );

        }


        return result;

    }


    function escapeDriveQuery(
        value
    ) {

        return String(
            value
        )
            .replace(
                /\\/g,
                "\\\\"
            )
            .replace(
                /'/g,
                "\\'"
            );

    }


    function escapeHtml(
        value
    ) {

        return String(
            value ?? ""
        )
            .replace(
                /[&<>'"]/g,
                character =>
                    ({

                        "&":
                            "&amp;",

                        "<":
                            "&lt;",

                        ">":
                            "&gt;",

                        "'":
                            "&#39;",

                        '"':
                            "&quot;"

                    }[
                        character
                    ])

            );

    }


    function escapeAttr(
        value
    ) {

        return escapeHtml(
            value
        );

    }


    function emptyRow(
        span,
        text
    ) {

        return `

            <tr>

                <td
                    colspan="${span}"
                    class="empty"
                >
                    ${escapeHtml(
                        text
                    )}
                </td>

            </tr>

        `;

    }


    function setAuthStatus(
        text,
        error = false
    ) {

        const element =
            $("auth-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setMovementStatus(
        text,
        error = false
    ) {

        const element =
            $("movement-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setInventoryManageStatus(
        text,
        error = false
    ) {

        const element =
            $("inventory-manage-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setSyncStatus(
        text,
        error = false
    ) {

        const element =
            $("sync-status");


        element.textContent =
            text;


        element.className =
            `muted ${
                error
                    ? "error"
                    : ""
            }`;

    }


})();