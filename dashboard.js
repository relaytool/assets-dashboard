(() => {

    "use strict";


    const SHEETS_API =
        "https://sheets.googleapis.com/v4/spreadsheets";

    const DRIVE_API =
        "https://www.googleapis.com/drive/v3/files";


    const SESSION_KEY = "fmAssetSession";
    let tokenRequestPromise = null;

    const state = {

        idTokenPayload: null,

        accessToken: null,

        workbookId: null,

        workbookName: null,

        inventory: [],

        transactions: [],

        todayLoads: [],

        clients: []

    };


    const $ =
        id => document.getElementById(id);


    document.addEventListener(
        "DOMContentLoaded",
        () => {

            bindEvents();

            setDefaultTimestamp();

            waitForGoogle();

        }
    );


    function bindEvents() {

        $("grant-access")
            .addEventListener(
                "click",
                requestSheetAccess
            );


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

        $("add-asset-row").addEventListener("click", () => addAssetRow());
        $("asset-rows").addEventListener("input", updateMovementPreview);
        $("asset-rows").addEventListener("change", updateMovementPreview);
        $("asset-rows").addEventListener("click", event => {
            const button = event.target.closest("[data-remove-asset]");
            if (!button) return;
            removeAssetRow(Number(button.dataset.removeAsset));
        });


        $("movement")
            .addEventListener(
                "change",
                handleMovementChange
            );


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


    function waitForGoogle() {

        let checks = 0;

        const maxChecks = 150;


        const timer =
            setInterval(
                () => {

                    checks++;


                    if (
                        window.google?.accounts?.id &&
                        window.google?.accounts?.oauth2
                    ) {

                        clearInterval(timer);

                        initializeGoogle();

                    }


                    if (
                        checks >= maxChecks
                    ) {

                        clearInterval(timer);

                        setAuthStatus(
                            "Google services could not be loaded. Check your internet connection.",
                            true
                        );

                    }

                },
                100
            );

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

        google.accounts.id.renderButton($("google-signin-button"), {
            theme: "outline", size: "large", text: "signin_with", shape: "rectangular", width: 280
        });

        const saved = readSavedSession();
        if (saved) {
            setUserProfile(saved);
            attemptSilentAccess(saved.email);
        }
        google.accounts.id.prompt();
    }


    function handleCredentialResponse(response) {
        try {
            state.idTokenPayload = decodeJwtPayload(response.credential);
            saveSession();
            setUserProfile(state.idTokenPayload);
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

        acquireAccessToken(silent ? "none" : "consent")
            .then(async () => {
                hideLogin();
                await loadDashboard();
            })
            .catch(error => {
                $("grant-access").classList.remove("hidden");
                if (silent) {
                    setAuthStatus("Google account restored. Connect Google Sheets to continue.");
                } else {
                    setAuthStatus(error.message || "Google authorization failed. Click Connect Google Sheets to try again.", true);
                }
            });
    }


    function attemptSilentAccess(email) {
        acquireAccessToken("none", email)
            .then(async () => {
                hideLogin();
                await loadDashboard();
            })
            .catch((error) => {
      console.warn("Silent Google Sheets/Drive authorization failed:", error);
                $("grant-access").classList.remove("hidden");
                setAuthStatus("Google account restored. Allow Sheets & Drive access to continue.");
            });
    }


    function acquireAccessToken(prompt = "none", email) {
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
            const timeoutId = setTimeout(() => finish(reject, new Error("Google's sign-in popup didn't open (it may have been blocked). Please allow popups for this site and try again.")), 20000);
            const tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: CONFIG.GOOGLE_CLIENT_ID,
                scope: CONFIG.OAUTH_SCOPES,
                callback: response => {
                    if (response.error) { finish(reject, new Error(`Google authorization failed: ${response.error}`)); return; }
                    state.accessToken = response.access_token;
                    finish(resolve, response.access_token);
                }
            });
            tokenClient.requestAccessToken({
                prompt,
                login_hint: email || state.idTokenPayload?.email || readSavedSession()?.email || undefined
            });
        });
        return tokenRequestPromise;
    }


    async function loadDashboard() {

        if (!state.accessToken) {
            return;
        }


        setSyncStatus(
            "Syncing with Google Sheets..."
        );


        try {

            // Main tracker workbook: READ ONLY. Tries the primary name
            // first, then falls back to the legacy "Copy of..." name.
            const workbook =
                await findMainWorkbook();


            if (!workbook) {

                throw new Error(
                    `Workbook "${CONFIG.PRIMARY_WORKBOOK_NAME}" (or "${CONFIG.FALLBACK_WORKBOOK_NAME}") was not found in your Google Drive.`
                );

            }


            state.workbookId =
                workbook.id;


            state.workbookName =
                workbook.name;


            $("workbook-name")
                .textContent =
                workbook.name;


            // Assets Inventory Ledger: the read/write source for
            // inventory + transactions. Opened directly by ID.
            const ledgerId =
                CONFIG.INVENTORY_LEDGER_SHEET_ID;


            const metadata =
                await sheetsGet(
                    `/${encodeURIComponent(
                        ledgerId
                    )}`
                );


            let sheetTitles =
                (metadata.sheets || [])
                    .map(
                        sheet =>
                            sheet.properties.title
                    );


            const missingSheets = [];


            if (
                !sheetTitles.includes(
                    CONFIG.INVENTORY_SHEET_NAME
                )
            ) {

                missingSheets.push(
                    CONFIG.INVENTORY_SHEET_NAME
                );

            }


            if (
                !sheetTitles.includes(
                    CONFIG.TRANSACTIONS_SHEET_NAME
                )
            ) {

                missingSheets.push(
                    CONFIG.TRANSACTIONS_SHEET_NAME
                );

            }


            if (missingSheets.length) {

                await createSheets(
                    missingSheets
                );

                sheetTitles =
                    sheetTitles.concat(
                        missingSheets
                    );

            }


            const [
                inventoryRows,
                transactionRows,
                clientListRows,
                mainRows
            ] = await Promise.all([
                getValues(
                    ledgerId,
                    CONFIG.INVENTORY_SHEET_NAME
                ),
                getValues(
                    ledgerId,
                    CONFIG.TRANSACTIONS_SHEET_NAME
                ),
                getValues(
                    ledgerId,
                    CONFIG.CLIENT_LIST_SHEET_NAME
                ),
                getValues(
                    state.workbookId,
                    CONFIG.MAIN_SHEET_NAME
                )
            ]);


            state.inventory =
                parseInventory(
                    inventoryRows
                );


            state.transactions =
                parseTransactions(
                    transactionRows
                );


            state.todayLoads =
                parseTodayLoads(
                    mainRows
                );


            state.clients =
                parseClients(
                    clientListRows
                );


            renderDashboard();


            setSyncStatus(
                `Synced at ${new Date().toLocaleTimeString()}`
            );


        } catch (error) {

            console.error(error);

            setSyncStatus(
                error.message ||
                "Unable to load spreadsheet.",
                true
            );

        }

    }


    async function findWorkbook(name) {

        const query = [

            `name = '${escapeDriveQuery(name)}'`,

            `mimeType = 'application/vnd.google-apps.spreadsheet'`,

            `trashed = false`

        ].join(" and ");


        const url =

            `${DRIVE_API}?q=${encodeURIComponent(query)}` +

            `&pageSize=20` +

            `&fields=files(id,name,mimeType,modifiedTime,webViewLink)`;


        const data =
            await fetchJson(
                url,
                {
                    headers:
                        authHeaders()
                }
            );


        return data.files?.[0] ||
            null;

    }


    // Looks for the primary tracker name first, falling back to the
    // legacy "Copy of..." name. Read-only - never written to.
    async function findMainWorkbook() {

        return (
            await findWorkbook(
                CONFIG.PRIMARY_WORKBOOK_NAME
            )
        ) || (
            await findWorkbook(
                CONFIG.FALLBACK_WORKBOOK_NAME
            )
        );

    }


    // Sheet-creation only ever targets the Assets Inventory Ledger -
    // the tracker workbook is never written to.
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
                        "User"
                    ]
                ]

            );

        }

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


    function parseTransactions(
        rows
    ) {

        if (!rows.length) {
            return [];
        }


        const header =
            rows[0].map(
                normalizeHeader
            );


        const idx = {

            timestamp:
                findColumn(
                    header,
                    [
                        "timestamp",
                        "date",
                        "datetime"
                    ]
                ),

            client:
                findColumn(
                    header,
                    [
                        "client",
                        "client name"
                    ]
                ),

            movement:
                findColumn(
                    header,
                    [
                        "movement",
                        "type",
                        "direction"
                    ]
                ),

            asset:
                findColumn(
                    header,
                    [
                        "asset",
                        "asset name",
                        "item"
                    ]
                ),

            quantity:
                findColumn(
                    header,
                    [
                        "quantity",
                        "qty"
                    ]
                ),

            user:
                findColumn(
                    header,
                    [
                        "user",
                        "entered by",
                        "email"
                    ]
                )

        };


        return rows

            .slice(1)

            .map(
                row => ({

                    timestamp:
                        idx.timestamp >= 0
                            ? row[idx.timestamp] ?? ""
                            : "",

                    client:
                        idx.client >= 0
                            ? row[idx.client] ?? ""
                            : "",

                    movement:
                        idx.movement >= 0
                            ? row[idx.movement] ?? ""
                            : "",

                    asset:
                        idx.asset >= 0
                            ? row[idx.asset] ?? ""
                            : "",

                    quantity:
                        idx.quantity >= 0
                            ? numericValue(
                                row[idx.quantity]
                            )
                            : 0,

                    user:
                        idx.user >= 0
                            ? row[idx.user] ?? ""
                            : ""

                })
            )

            .filter(
                item =>
                    item.asset ||
                    item.client
            );

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

        if (!rows.length) {
            return [];
        }

        const header =
            rows[0].map(normalizeHeader);

        const clientIdx =
            findColumn(
                header,
                [
                    "client name",
                    "client",
                    "name"
                ]
            );

        if (clientIdx < 0) {
            return [];
        }

        const clients = new Set();

        rows.slice(1).forEach(row => {
            const value =
                String(row[clientIdx] ?? "").trim();

            if (value) {
                clients.add(value);
            }
        });

        return [...clients].sort(
            (a, b) => a.localeCompare(b)
        );
    }


    function parseTodayLoads(
        rows
    ) {

        if (!rows.length) {
            return [];
        }


        const headerRowIndex =
            findMainHeaderRow(
                rows
            );


        if (headerRowIndex < 0) {
            return [];
        }


        const headers =
            rows[
                headerRowIndex
            ].map(
                normalizeHeader
            );


        const clientIdx =
            findColumn(
                headers,
                [
                    "cleint",
                    "client",
                    "collection client"
                ]
            );


        const loadTypeIdx =
            findColumn(
                headers,
                [
                    "load type",
                    "loadtype"
                ]
            );


        const plannedArrivalIdx =
            findColumn(
                headers,
                [
                    "planned arrival",
                    "plannedarrival"
                ]
            );


        const palletsIdx =
            findColumn(
                headers,
                [
                    "pallets",
                    "pallet"
                ]
            );


        const looseIdx =
            findColumn(
                headers,
                [
                    "bags / loose parcles",
                    "bags / loose parcels",
                    "bags loose parcles",
                    "bags loose parcels"
                ]
            );


        const locationIdx =
            findColumn(
                headers,
                [
                    "/",
                    "location",
                    "site"
                ]
            );


        if (
            clientIdx < 0 ||
            loadTypeIdx < 0 ||
            plannedArrivalIdx < 0
        ) {

            return [];

        }


        const today =
            new Date();


        const todayDay =
            today.getDate();


        const todayMonth =
            today.getMonth() + 1;


        const todayYear =
            today.getFullYear();


        let activeDate =
            null;


        const loads = [];


        for (
            let i =
                headerRowIndex + 1;

            i < rows.length;

            i++
        ) {

            const row =
                rows[i] || [];


            let sectionDate =
                null;


            for (
                let col = 0;

                col < Math.min(
                    4,
                    row.length
                );

                col++
            ) {

                sectionDate =
                    parseTrackerDate(
                        row[col]
                    );


                if (sectionDate) {
                    break;
                }

            }


            if (sectionDate) {

                activeDate =
                    sectionDate;

                continue;

            }


            const client =
                String(
                    row[clientIdx] ??
                    ""
                ).trim();


            const loadType =
                String(
                    row[loadTypeIdx] ??
                    ""
                ).trim();


            const plannedArrival =
                String(
                    row[plannedArrivalIdx] ??
                    ""
                ).trim();


            const location =
                locationIdx >= 0
                    ? String(
                        row[locationIdx] ??
                        ""
                    ).trim()
                    : "";


            const pallets =
                palletsIdx >= 0
                    ? String(
                        row[palletsIdx] ??
                        ""
                    ).trim()
                    : "";


            const loose =
                looseIdx >= 0
                    ? String(
                        row[looseIdx] ??
                        ""
                    ).trim()
                    : "";


            if (
                !client &&
                !loadType &&
                !plannedArrival
            ) {
                continue;
            }


            if (
                client.toLowerCase() ===
                    "cleint" ||

                client.toLowerCase() ===
                    "client"
            ) {
                continue;
            }


            if (
                location &&
                location.toUpperCase() !==
                    CONFIG.LOCATION.toUpperCase()
            ) {
                continue;
            }


            if (!activeDate) {
                continue;
            }


            if (
                activeDate.day !==
                    todayDay ||

                activeDate.month !==
                    todayMonth ||

                activeDate.year !==
                    todayYear
            ) {
                continue;
            }


            const direction =
                loadType.toUpperCase() ===
                    "RETURNS"

                    ? "Outbound"

                    : "Inbound";


            loads.push({

                client,

                loadType,

                plannedArrival,

                pallets,

                loose,

                direction

            });

        }


        return loads;

    }


    function parseTrackerDate(
        value
    ) {

        if (
            value === null ||
            value === undefined
        ) {
            return null;
        }


        const text =
            String(value).trim();


        if (!text) {
            return null;
        }


        const shortDate =
            text.match(
                /^(\d{1,2})[\/\-](\d{1,2})$/
            );


        if (shortDate) {

            return {

                day:
                    Number(
                        shortDate[1]
                    ),

                month:
                    Number(
                        shortDate[2]
                    ),

                year:
                    new Date()
                        .getFullYear()

            };

        }


        const parsed =
            new Date(text);


        if (
            !Number.isNaN(
                parsed.getTime()
            )
        ) {

            return {

                day:
                    parsed.getDate(),

                month:
                    parsed.getMonth() + 1,

                year:
                    parsed.getFullYear()

            };

        }


        return null;

    }


    function renderAudit() {

        const rows = state.transactions.slice(-30).reverse();

        const count = $("audit-count");
        if (count) count.textContent = rows.length.toLocaleString();

        const body = $("transactions-body");
        if (!body) return;

        body.innerHTML = rows.length
            ? rows.map(item => `
                <tr>
                    <td>${escapeHtml(formatTimestamp(item.timestamp))}</td>
                    <td><strong>${escapeHtml(item.client || "")}</strong></td>
                    <td><span class="movement-badge ${movementClass(item.movement)}">${escapeHtml(item.movement || "")}</span></td>
                    <td>${escapeHtml(item.asset || "")}</td>
                    <td class="num">${formatNumber(item.quantity)}</td>
                    <td>${escapeHtml(item.user || "")}</td>
                </tr>
            `).join("")
            : emptyRow(6, "No asset movements recorded yet.");
    }


    /* =====================================================
       DASHBOARD
       ===================================================== */


    function renderDashboard() {

        const inbound =
            state.todayLoads.filter(
                item =>
                    item.direction ===
                    "Inbound"
            );


        const outbound =
            state.todayLoads.filter(
                item =>
                    item.direction ===
                    "Outbound"
            );


        $("inbound-badge")
            .textContent =
            inbound.length
                .toLocaleString();


        $("outbound-badge")
            .textContent =
            outbound.length
                .toLocaleString();


        /*
         * NEW:
         * Warehouse asset cards
         */

        renderWarehouseAssets();


        renderClientCards();


        renderLoadTable(
            $("inbound-body"),
            inbound
        );


        renderLoadTable(
            $("outbound-body"),
            outbound
        );


        /*
         * Multi-asset movement form
         */
        const clients = state.clients.filter(client => client !== "HSC London(Self)");
        $("client").innerHTML = clients.length ? `<option value="">Select client</option>${clients.map(client => `<option value="${escapeAttr(client)}">${escapeHtml(client)}</option>`).join("")}` : `<option value="">No clients found</option>`;
        const existingRows = [...document.querySelectorAll(".asset-row")];
        if (!existingRows.length) addAssetRow(false);
        document.querySelectorAll(".asset-select").forEach(select => { const value = select.value; select.innerHTML = assetOptions(value); });
        updateAssetRows();
        handleMovementChange();
        updateMovementPreview();


        renderAudit();

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


        state.transactions

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


                    if (
                        movement ===
                        "RECEIVED"
                    ) {

                        entry.net +=
                            quantity;

                        entry.received +=
                            quantity;

                    }


                    if (
                        movement ===
                        "SENT"
                    ) {

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
    
        if (!state.inventory.length) {
            container.innerHTML = `
                <div class="warehouse-empty">
                    No assets configured.
                </div>
            `;
            return;
        }
    
        container.innerHTML = state.inventory.map(item => {
    
            const assetName = String(item.asset).trim();
            const balance = Number(item.balance) || 0;
    
            /*
             * Calculate today's movement for this asset.
             *
             * RECEIVED = positive
             * SENT     = negative
             * DISCARD  = negative
             */
            const movementTotal = state.transactions
                .filter(transaction => {
                    return (
                        String(transaction.asset).trim().toLowerCase() ===
                        assetName.toLowerCase()
                        &&
                        isTodayTransaction(transaction.timestamp)
                    );
                })
                .reduce((total, transaction) => {
    
                    const movement =
                        String(transaction.movement)
                            .trim()
                            .toUpperCase();
    
                    const quantity =
                        Number(transaction.quantity) || 0;
    
                    if (movement === "RECEIVED") {
                        return total + quantity;
                    }
    
                    if (
                        movement === "SENT" ||
                        movement === "DISCARD"
                    ) {
                        return total - quantity;
                    }
    
                    return total;
                }, 0);
    
            let indicator = "";
    
            if (movementTotal > 0) {
                indicator = `
                    <span class="asset-movement positive">
                        +${formatNumber(movementTotal)}
                    </span>
                `;
            } else if (movementTotal < 0) {
                indicator = `
                    <span class="asset-movement negative">
                        ${formatNumber(movementTotal)}
                    </span>
                `;
            } else {
                indicator = `
                    <span class="asset-movement neutral">
                        0
                    </span>
                `;
            }
    
            return `
                <div class="warehouse-asset-card">
    
                    <div class="warehouse-asset-name">
                        ${escapeHtml(assetName)}
                    </div>
    
                    <div class="warehouse-asset-count">
                        ${formatNumber(balance)}
                    </div>
    
                    <div class="warehouse-asset-movement">
                        ${indicator}
                        <span class="movement-label">today</span>
                    </div>
    
                </div>
            `;
        }).join("");
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
        const query =
            String($("client-search")?.value || "")
                .trim()
                .toLowerCase();

        const clients =
            state.clients
                .filter(
                    client =>
                        client.toLowerCase().includes(query)
                )
                .sort(
                    (a, b) => a.localeCompare(b)
                );

        container.innerHTML =
            clients.length
                ? clients.map(
                    client => `
                        <article
                            class="client-card client-ledger-card"
                            data-client="${escapeAttr(client)}"
                            tabindex="0"
                            role="button"
                        >
                            <div class="client-card-top">
                                <div>
                                    <div class="eyebrow">CLIENT</div>
                                    <h3>${escapeHtml(client)}</h3>
                                </div>
                            </div>

                            <div class="client-card-footer client-ledger-footer">
                                <span>Asset inventory</span>

                                <button
                                    type="button"
                                    class="secondary small-button client-show-more"
                                    data-client="${escapeAttr(client)}"
                                >
                                    Show more
                                </button>
                            </div>
                        </article>
                    `
                ).join("")
                : `
                    <div class="client-cards-empty">
                        <strong>
                            ${query ? "No clients found" : "No clients in Client List"}
                        </strong>
                        <span>
                            ${query
                                ? "Try another client name."
                                : "Add clients to the Client List sheet in the Assets Inventory Ledger workbook."}
                        </span>
                    </div>
                  `;
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

        const transactions =
            state.transactions.filter(
                item =>
                    String(item.client).trim().toLowerCase() ===
                    String(client).trim().toLowerCase()
            );

        const assets =
            getClientAssetSummary(transactions);

        $("client-details-title").textContent = client;

        $("client-details-subtitle").textContent =
            `${assets.length} asset type${assets.length === 1 ? "" : "s"} · ` +
            `${transactions.length} transaction${transactions.length === 1 ? "" : "s"}`;

        $("client-details-summary").innerHTML =
            assets.length
                ? assets.map(
                    item => `
                        <button
                            type="button"
                            class="client-asset-summary-row ${
                                item.net < 0
                                    ? "asset-net-negative"
                                    : item.net > 0
                                        ? "asset-net-positive"
                                        : "asset-net-zero"
                            }"
                            data-client-asset="${escapeAttr(item.asset)}"
                            title="Show ${escapeAttr(item.asset)} daily sent graph"
                        >
                            <span class="client-asset-name">
                                ${escapeHtml(item.asset)}
                            </span>

                            <span class="client-asset-quantity">
                                ${formatSignedNumber(item.net)}
                            </span>

                            <span class="client-asset-meta">
                                Sent − Received
                            </span>
                        </button>
                    `
                ).join("")
                : `
                    <div class="client-assets-empty">
                        <strong>No asset movements for this client</strong>
                        <span>
                            Transactions for this client will appear here automatically.
                        </span>
                    </div>
                  `;

        const firstAsset =
            assets[0]?.asset || "";

        renderClientAssetGraph(
            client,
            firstAsset
        );

        $("client-details-transactions").innerHTML =
            transactions.length
                ? transactions
                    .slice()
                    .sort(
                        (a, b) =>
                            new Date(b.timestamp) -
                            new Date(a.timestamp)
                    )
                    .map(
                        item => `
                            <tr>
                                <td>
                                    ${formatTransactionTime(item.timestamp)}
                                </td>

                                <td>
                                    <span
                                        class="movement-badge ${movementClass(item.movement)}"
                                    >
                                        ${escapeHtml(
                                            String(item.movement)
                                        )}
                                    </span>
                                </td>

                                <td>
                                    <strong>
                                        ${escapeHtml(
                                            String(item.asset)
                                        )}
                                    </strong>
                                </td>

                                <td class="num">
                                    ${formatNumber(item.quantity)}
                                </td>

                                <td>
                                    ${escapeHtml(
                                        String(item.user || "")
                                    )}
                                </td>
                            </tr>
                        `
                    )
                    .join("")
                : emptyRow(
                    5,
                    "No asset transactions recorded for this client."
                );

        $("client-details-modal")
            .classList
            .remove("hidden");
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
                        received: 0
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
        });

        return [...map.values()]
            .map(
                item => ({
                    ...item,
                    net:
                        item.sent -
                        item.received
                })
            )
            .sort(
                (a, b) =>
                    a.asset.localeCompare(b.asset)
            );
    }


    function renderClientAssetGraph(
        client,
        asset
    ) {

        const title =
            $("client-graph-title");

        const subtitle =
            $("client-graph-subtitle");

        const graph =
            $("client-sent-graph");

        if (!asset) {

            title.textContent =
                "Daily sent vs received";

            subtitle.textContent =
                "Select an asset type above to see its daily sent and received history.";

            graph.innerHTML =
                `<div class="client-graph-empty">
                    No asset transaction history to graph.
                 </div>`;

            return;
        }

        title.textContent =
            `${asset} · sent vs received daily`;

        subtitle.textContent =
            "Daily quantity sent and received for this asset, from the Asset Transactions ledger.";

        const rows =
            state.transactions.filter(
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

        const points =
            [...daily.entries()]
                .sort((a, b) => a[0].localeCompare(b[0]));

        if (!points.length) {

            graph.innerHTML =
                `<div class="client-graph-empty">
                    No sent or received transactions recorded for ${escapeHtml(asset)}.
                 </div>`;

            return;
        }

        const width =
            Math.max(
                680,
                points.length * 82
            );

        const height = 300;
        const left = 56;
        const right = 24;
        const top = 42;
        const bottom = 58;
        const chartWidth = width - left - right;
        const chartHeight = height - top - bottom;

        const maxValue =
            Math.max(
                ...points.flatMap(([, value]) => [value.sent, value.received]),
                1
            );

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

        const xLabels = points.map(
            ([date], index) => `
                <text
                    x="${xFor(index).toFixed(1)}"
                    y="${height - 20}"
                    text-anchor="middle"
                    class="graph-label"
                >
                    ${escapeHtml(formatGraphDate(date))}
                </text>
            `
        ).join("");

        const pointsMarkup = points.map(
            ([date, value], index) => {
                const x = xFor(index);
                const label = formatGraphDate(date);

                return `
                    <g>
                        <title>
                            ${escapeHtml(label)}: ${formatNumber(value.sent)} sent, ${formatNumber(value.received)} received
                        </title>
                        <circle
                            cx="${x.toFixed(1)}"
                            cy="${yFor(value.sent).toFixed(1)}"
                            r="4"
                            class="graph-point-sent"
                        ></circle>
                        <circle
                            cx="${x.toFixed(1)}"
                            cy="${yFor(value.received).toFixed(1)}"
                            r="4"
                            class="graph-point-received"
                        ></circle>
                    </g>
                `;
            }
        ).join("");

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
                    viewBox="0 0 ${width} ${height}"
                    role="img"
                    aria-label="Daily sent versus received line graph for ${escapeHtml(asset)}"
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


    function renderLoadTable(
        body,
        items
    ) {

        body.innerHTML =

            items.length

                ? items.map(
                    item => `

                        <tr>

                            <td>

                                <strong>
                                    ${escapeHtml(
                                        item.client ||
                                        "Unknown"
                                    )}
                                </strong>

                                ${
                                    item.loadType

                                        ? `

                                            <span
                                                class="subtext"
                                            >
                                                ${escapeHtml(
                                                    item.loadType
                                                )}
                                            </span>

                                          `

                                        : ""
                                }

                            </td>

                            <td>
                                ${escapeHtml(
                                    item.plannedArrival ||
                                    "Not planned"
                                )}
                            </td>

                            <td class="num">
                                ${escapeHtml(
                                    item.pallets ||
                                    "0"
                                )}
                            </td>

                            <td>
                                ${escapeHtml(
                                    item.loose ||
                                    "0"
                                )}
                            </td>

                        </tr>

                    `
                ).join("")

                : emptyRow(
                    4,
                    "No matching loads found for today."
                );

    }


    async function recordMovement(event) {
        event.preventDefault();
        const movement = $("movement").value;
        const client = movement === "DISCARD" ? "HSC London (Self)" : $("client").value.trim();
        const entries = getAssetEntries();
        const user = state.idTokenPayload?.email || readSavedSession()?.email || state.idTokenPayload?.name || "Google user";

        if (!client || !entries.length || entries.some(x => !x.asset || !Number.isInteger(x.quantity) || x.quantity <= 0)) {
            setMovementStatus("Select a client and add at least one asset with a whole quantity greater than zero.", true);
            return;
        }

        const seen = new Set();
        const balances = new Map(state.inventory.map(item => [item.asset.toLowerCase(), item.balance]));
        const prepared = [];
        for (const entry of entries) {
            const key = entry.asset.toLowerCase();
            if (seen.has(key)) { setMovementStatus(`You have selected ${entry.asset} more than once. Combine the quantities.`, true); return; }
            seen.add(key);
            const item = state.inventory.find(x => x.asset.toLowerCase() === key);
            if (!item) { setMovementStatus(`Asset "${entry.asset}" is not present in Inventory.`, true); return; }
            const current = balances.get(key) || 0;
            const next = movement === "RECEIVED" ? current + entry.quantity : current - entry.quantity;
            if (movement !== "RECEIVED" && next < 0) { setMovementStatus(`Cannot remove ${entry.quantity} ${item.asset}. Current balance is ${formatNumber(current)}.`, true); return; }
            balances.set(key, next);
            prepared.push({ ...entry, item, newBalance: next });
        }

        try {
            setMovementStatus("Recording movements...");
            const ledgerId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
            const transactionRange = `${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A:F`;
            const timestamp = new Date().toISOString();
            await sheetsPost(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(transactionRange)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`, {
                values: prepared.map(x => [timestamp, client, movement, x.asset, x.quantity, user])
            });
            for (const entry of prepared) {
                const balanceCell = columnLetter(entry.item.balanceColumn) + entry.item.rowNumber;
                const inventoryRange = `${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!${balanceCell}`;
                await sheetsPut(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(inventoryRange)}?valueInputOption=USER_ENTERED`, { range: inventoryRange, majorDimension: "ROWS", values: [[entry.newBalance]] });
            }
            resetMovementForm();
            setMovementStatus(`${prepared.length} movement${prepared.length === 1 ? "" : "s"} recorded successfully.`);
            await loadDashboard();
        } catch (error) {
            console.error(error);
            setMovementStatus(error.message || "Unable to record movement.", true);
        }
    }


    function assetOptions(selected = "") {
        return state.inventory.length ? `<option value="">Select asset type</option>${state.inventory.map(item => `<option value="${escapeAttr(item.asset)}" ${item.asset === selected ? "selected" : ""}>${escapeHtml(item.asset)}</option>`).join("")}` : `<option value="">No assets configured</option>`;
    }

    function addAssetRow(focus = true) {
        const wrap = $("asset-rows");
        const row = document.createElement("div");
        row.className = "asset-row";
        row.innerHTML = `<div class="asset-row-number"></div><label><span>Asset type</span><select class="asset-select" required>${assetOptions()}</select></label><label><span>Quantity</span><input class="asset-quantity" type="number" min="1" step="1" inputmode="numeric" placeholder="0" required></label><button type="button" class="secondary remove-asset" data-remove-asset="0">Remove</button>`;
        wrap.appendChild(row);
        updateAssetRows();
        if (focus) row.querySelector(".asset-select").focus();
    }

    function removeAssetRow(index) {
        const rows = [...document.querySelectorAll(".asset-row")];
        if (rows.length <= 1) return;
        rows[index]?.remove();
        updateAssetRows();
        updateMovementPreview();
    }

    function updateAssetRows() {
        const rows = [...document.querySelectorAll(".asset-row")];
        rows.forEach((row, index) => { row.querySelector(".asset-row-number").textContent = String(index + 1).padStart(2, "0"); row.querySelector("[data-remove-asset]").dataset.removeAsset = index; row.querySelector("[data-remove-asset]").disabled = rows.length === 1; });
    }

    function getAssetEntries() {
        return [...document.querySelectorAll(".asset-row")].map(row => ({ asset: row.querySelector(".asset-select").value, quantity: Number(row.querySelector(".asset-quantity").value) }));
    }

    function updateMovementPreview() {
        const client = $("movement").value === "DISCARD" ? "HSC London (Self)" : $("client").value.trim();
        const entries = getAssetEntries().filter(x => x.asset && Number.isInteger(x.quantity) && x.quantity > 0);
        if (!client || !entries.length) { $("preview-text").textContent = "Select a client and add one or more asset types with quantities."; return; }
        $("preview-text").textContent = `${movementLabel($("movement").value)} ${entries.map(x => `${formatNumber(x.quantity)} × ${x.asset}`).join(" • ")} for ${client}`;
    }

    function resetMovementForm() { $("movement-form").reset(); $("client").disabled = false; $("asset-rows").innerHTML = ""; addAssetRow(false); setDefaultTimestamp(); handleMovementChange(); }


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


            const range =
                `${quoteSheetName(
                    CONFIG.INVENTORY_SHEET_NAME
                )}!A:B`;


            await sheetsPost(

                `/${encodeURIComponent(
                    CONFIG.INVENTORY_LEDGER_SHEET_ID
                )}/values/${encodeURIComponent(
                    range
                )}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,

                {

                    values: [[

                        assetName,

                        balance

                    ]]

                }

            );


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


            const assetCell =
                columnLetter(
                    item.assetColumn
                ) +
                item.rowNumber;


            const range =
                `${quoteSheetName(
                    CONFIG.INVENTORY_SHEET_NAME
                )}!${assetCell}`;


            await sheetsPut(

                `/${encodeURIComponent(
                    CONFIG.INVENTORY_LEDGER_SHEET_ID
                )}/values/${encodeURIComponent(
                    range
                )}?valueInputOption=USER_ENTERED`,

                {

                    range,

                    majorDimension:
                        "ROWS",

                    values: [[
                        cleanName
                    ]]

                }

            );


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

                            <td class="num">

                                ${formatNumber(
                                    item.balance
                                )}

                            </td>

                            <td
                                class="actions-cell"
                            >

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

                : emptyRow(
                    3,
                    "No asset types configured."
                );

    }


    function handleMovementChange() {

        const movement =
            $("movement").value;


        const client =
            $("client");


        if (
            movement ===
            "DISCARD"
        ) {

            client.value =
                "HSC London (Self)";


            client.disabled =
                true;

        } else {

            client.disabled =
                false;


            if (
                client.value ===
                "HSC London (Self)"
            ) {

                client.value =
                    "";

            }

        }

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


    function signOut() {
        const saved = readSavedSession();
        if (state.idTokenPayload?.sub) { try { google.accounts.id.revoke(state.idTokenPayload.sub, () => {}); } catch (_) {} }
        if (saved?.sub && saved.sub !== state.idTokenPayload?.sub) { try { google.accounts.id.revoke(saved.sub, () => {}); } catch (_) {} }
        localStorage.removeItem(SESSION_KEY);
        state.idTokenPayload = null; state.accessToken = null; state.workbookId = null; state.workbookName = null; state.inventory = []; state.transactions = []; state.todayLoads = []; state.clients = [];
        $("dashboard").classList.add("hidden"); $("login-card").classList.remove("hidden"); $("google-signin-button").classList.remove("hidden"); $("grant-access").classList.add("hidden"); $("sign-out").classList.add("hidden"); $("user-photo").classList.add("hidden"); $("user-name").textContent = "Not signed in"; $("user-email").textContent = "";
        closeInventoryManager();
    }


    function saveSession() { if (state.idTokenPayload) localStorage.setItem(SESSION_KEY, JSON.stringify({ name: state.idTokenPayload.name || "Google user", email: state.idTokenPayload.email || "", picture: state.idTokenPayload.picture || "", sub: state.idTokenPayload.sub || "" })); }
    function readSavedSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch (_) { return null; } }
    function setUserProfile(profile) { $("user-name").textContent = profile?.name || "Google user"; $("user-email").textContent = profile?.email || ""; if (profile?.picture) { $("user-photo").src = profile.picture; $("user-photo").classList.remove("hidden"); } }
    function hideLogin() { $("google-signin-button").classList.add("hidden"); $("grant-access").classList.add("hidden"); $("sign-out").classList.remove("hidden"); $("login-card").classList.add("hidden"); $("dashboard").classList.remove("hidden"); }


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


    async function fetchJson(url, options = {}) {
        let response = await fetch(url, options);
        if (response.status === 401 && !options.__retried) {
            try {
                await acquireAccessToken("none");
                return fetchJson(url, { ...options, __retried: true, headers: { ...(options.headers || {}), ...authHeaders() } });
            } catch (_) {
                throw new Error("Google Sheets access expired. Please reconnect Google Sheets.");
            }
        }
        const text = await response.text(); let data = {}; try { data = text ? JSON.parse(text) : {}; } catch (_) {}
        if (!response.ok) throw new Error(data?.error?.message || `Request failed (${response.status})`);
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