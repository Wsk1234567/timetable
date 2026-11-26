(function () {
  "use strict";

  const META_KEY = "my-timetable-sync-v1";
  const APP_PROPERTY_KEY = "timetableApp";
  const APP_PROPERTY_VALUE = "wsk-timetable-v1";
  const SHEET_NAME = "Activity Log";
  const FILE_NAME = "我的时间表资料";
  const HEADERS = [
    "revision_id",
    "event_id",
    "series_id",
    "title",
    "date",
    "start_time",
    "end_time",
    "location",
    "action",
    "updated_at",
    "device_id",
  ];
  const SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
  ].join(" ");

  const uid = () =>
    globalThis.crypto?.randomUUID?.() ||
    `${Date.now()}-${Math.random().toString(36).slice(2)}`;

  function loadMeta() {
    let saved = {};
    try {
      saved = JSON.parse(localStorage.getItem(META_KEY) || "{}");
    } catch {
      saved = {};
    }
    return {
      sheetId: saved.sheetId || "",
      deviceId: saved.deviceId || uid(),
      pending: Array.isArray(saved.pending) ? saved.pending : [],
      migrated: Boolean(saved.migrated),
      lastSyncAt: saved.lastSyncAt || "",
    };
  }

  function revisionFor(activity, action, deviceId) {
    return [
      uid(),
      activity.id,
      activity.seriesId || "",
      action === "delete" ? "" : activity.title || "",
      action === "delete" ? "" : activity.date || "",
      action === "delete" ? "" : activity.start || "",
      action === "delete" ? "" : activity.end || "",
      action === "delete" ? "" : activity.location || "",
      action,
      new Date().toISOString(),
      deviceId,
    ];
  }

  function reduceRows(rows) {
    const current = new Map();
    for (const row of rows.slice(1)) {
      if (!row?.[1] || !["upsert", "delete"].includes(row[8])) continue;
      current.set(row[1], row);
    }
    return [...current.values()]
      .filter((row) => row[8] === "upsert")
      .map((row) => ({
        id: row[1],
        ...(row[2] ? { seriesId: row[2] } : {}),
        title: row[3] || "未命名活动",
        date: row[4] || "",
        start: row[5] || "09:00",
        end: row[6] || "10:00",
        location: row[7] || "",
      }))
      .filter((activity) => /^\d{4}-\d{2}-\d{2}$/.test(activity.date))
      .sort(
        (a, b) =>
          a.date.localeCompare(b.date) ||
          a.start.localeCompare(b.start) ||
          a.title.localeCompare(b.title),
      );
  }

  function create({ getActivities, setActivities, onStatus }) {
    const config = window.TIMETABLE_CONFIG || {};
    const meta = loadMeta();
    let token = "";
    let tokenClient = null;
    let syncing = false;
    let expired = false;

    function saveMeta() {
      localStorage.setItem(META_KEY, JSON.stringify(meta));
    }

    function report(state, message = "") {
      onStatus?.({
        state,
        message,
        pending: meta.pending.length,
        sheetId: meta.sheetId,
        lastSyncAt: meta.lastSyncAt,
      });
    }

    async function api(url, options = {}) {
      if (!token) throw new Error("AUTH_REQUIRED");
      const response = await fetch(url, {
        ...options,
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          ...(options.headers || {}),
        },
      });
      if (response.status === 401) {
        token = "";
        expired = true;
        report("reconnect", "Google 授权已到期，请重新连接");
        throw new Error("AUTH_REQUIRED");
      }
      if (!response.ok) {
        const detail = await response.json().catch(() => ({}));
        const error = new Error(
          detail?.error?.message || `Google API ${response.status}`,
        );
        error.status = response.status;
        throw error;
      }
      return response.status === 204 ? null : response.json();
    }

    async function validateSheet(sheetId) {
      try {
        await api(
          `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(sheetId)}?fields=id,trashed`,
        );
        return true;
      } catch (error) {
        if (error.status === 404) return false;
        throw error;
      }
    }

    async function initializeSheet(sheetId) {
      const spreadsheet = await api(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties(sheetId,title)`,
      );
      const sheets = spreadsheet.sheets || [];
      const logSheet =
        sheets.find((sheet) => sheet.properties.title === SHEET_NAME) ||
        sheets[0];
      if (!logSheet) throw new Error("无法初始化 Google Sheet");
      const requests = [];
      if (logSheet.properties.title !== SHEET_NAME) {
        requests.push({
          updateSheetProperties: {
            properties: {
              sheetId: logSheet.properties.sheetId,
              title: SHEET_NAME,
              gridProperties: { frozenRowCount: 1 },
            },
            fields: "title,gridProperties.frozenRowCount",
          },
        });
      } else {
        requests.push({
          updateSheetProperties: {
            properties: {
              sheetId: logSheet.properties.sheetId,
              gridProperties: { frozenRowCount: 1 },
            },
            fields: "gridProperties.frozenRowCount",
          },
        });
      }
      requests.push({
        repeatCell: {
          range: {
            sheetId: logSheet.properties.sheetId,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: HEADERS.length,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.15, green: 0.21, blue: 0.36 },
              textFormat: {
                foregroundColor: { red: 1, green: 1, blue: 1 },
                bold: true,
              },
            },
            note: "由我的时间表自动维护，请在网页中修改活动。",
          },
          fields: "userEnteredFormat(backgroundColor,textFormat),note",
        },
      });
      const hasInfoSheet = sheets.some(
        (sheet) => sheet.properties.title === "说明",
      );
      if (!hasInfoSheet) {
        requests.push({
          addSheet: {
            properties: {
              title: "说明",
              gridProperties: { rowCount: 20, columnCount: 2 },
            },
          },
        });
      }
      await api(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`,
        {
          method: "POST",
          body: JSON.stringify({ requests }),
        },
      );
      await api(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(`'${SHEET_NAME}'!A1:K1`)}?valueInputOption=RAW`,
        {
          method: "PUT",
          body: JSON.stringify({ values: [HEADERS] }),
        },
      );
      await api(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent("'说明'!A1:A5")}?valueInputOption=RAW`,
        {
          method: "PUT",
          body: JSON.stringify({
            values: [
              ["我的时间表云端资料"],
              ["这份文件由时间表网站自动建立。"],
              ["请在时间表网页新增、编辑或删除活动。"],
              ["Activity Log 是同步记录，请勿手动修改表头或删除工作表。"],
              ["删除整份文件后，网站会在下次同步时自动重建。"],
            ],
          }),
        },
      );
    }

    async function createSheet() {
      const file = await api(
        "https://www.googleapis.com/drive/v3/files?fields=id,webViewLink",
        {
          method: "POST",
          body: JSON.stringify({
            name: FILE_NAME,
            mimeType: "application/vnd.google-apps.spreadsheet",
            appProperties: { [APP_PROPERTY_KEY]: APP_PROPERTY_VALUE },
          }),
        },
      );
      meta.sheetId = file.id;
      saveMeta();
      await initializeSheet(file.id);
      return file.id;
    }

    async function findOrCreateSheet() {
      if (meta.sheetId && (await validateSheet(meta.sheetId))) {
        await initializeSheet(meta.sheetId);
        return meta.sheetId;
      }
      if (meta.sheetId) meta.migrated = false;
      meta.sheetId = "";
      const query = [
        `appProperties has { key='${APP_PROPERTY_KEY}' and value='${APP_PROPERTY_VALUE}' }`,
        "mimeType='application/vnd.google-apps.spreadsheet'",
        "trashed=false",
      ].join(" and ");
      const result = await api(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&orderBy=createdTime&fields=files(id,name,createdTime)`,
      );
      meta.sheetId = result.files?.[0]?.id || "";
      saveMeta();
      if (!meta.sheetId) return createSheet();
      await initializeSheet(meta.sheetId);
      return meta.sheetId;
    }

    async function readRows() {
      const range = encodeURIComponent(`'${SHEET_NAME}'!A:K`);
      const data = await api(
        `https://sheets.googleapis.com/v4/spreadsheets/${meta.sheetId}/values/${range}?majorDimension=ROWS`,
      );
      return data.values || [HEADERS];
    }

    async function appendPending() {
      if (!meta.pending.length) return;
      const pending = [...meta.pending];
      const range = encodeURIComponent(`'${SHEET_NAME}'!A:K`);
      await api(
        `https://sheets.googleapis.com/v4/spreadsheets/${meta.sheetId}/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
        {
          method: "POST",
          body: JSON.stringify({ values: pending }),
        },
      );
      meta.pending.splice(0, pending.length);
      saveMeta();
    }

    function queue(upserts = [], deletes = []) {
      meta.pending.push(
        ...deletes.map((activity) =>
          revisionFor(activity, "delete", meta.deviceId),
        ),
        ...upserts.map((activity) =>
          revisionFor(activity, "upsert", meta.deviceId),
        ),
      );
      saveMeta();
      report(token ? "queued" : expired ? "reconnect" : "local");
    }

    async function sync() {
      if (syncing) return;
      if (!navigator.onLine) {
        report("offline", "离线中，活动已保存在此设备");
        return;
      }
      if (!token) {
        report(expired ? "reconnect" : "local");
        return;
      }
      syncing = true;
      report("syncing");
      try {
        await findOrCreateSheet();
        let rows = await readRows();
        if (!meta.migrated) {
          const cloudIds = new Set(
            rows
              .slice(1)
              .map((row) => row?.[1])
              .filter(Boolean),
          );
          const missing = getActivities().filter(
            (activity) => !cloudIds.has(activity.id),
          );
          queue(missing, []);
          meta.migrated = true;
          saveMeta();
        }
        await appendPending();
        rows = await readRows();
        setActivities(reduceRows(rows));
        meta.lastSyncAt = new Date().toISOString();
        saveMeta();
        report("synced");
      } catch (error) {
        if (error.message !== "AUTH_REQUIRED") {
          report("error", error.message || "同步失败，请稍后重试");
        }
      } finally {
        syncing = false;
      }
    }

    function ensureTokenClient() {
      const clientId = config.googleClientId || "";
      if (!clientId) {
        report("unconfigured", "尚未设置 Google Client ID");
        return false;
      }
      if (!window.google?.accounts?.oauth2) {
        report("error", "Google 登录载入失败，请检查网络后重试");
        return false;
      }
      if (!tokenClient) {
        tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: clientId,
          scope: SCOPES,
          callback: async (response) => {
            if (response.error) {
              report(
                "error",
                response.error_description || "未完成 Google 授权",
              );
              return;
            }
            token = response.access_token;
            expired = false;
            const expiresIn = Math.max(
              60,
              Number(response.expires_in || 3600) - 60,
            );
            window.setTimeout(() => {
              if (token === response.access_token) {
                token = "";
                expired = true;
                report("reconnect", "授权已到期，活动仍保存在此设备");
              }
            }, expiresIn * 1000);
            await sync();
          },
          error_callback: () => report("error", "Google 登录窗口已关闭"),
        });
      }
      return true;
    }

    function connect() {
      if (!ensureTokenClient()) return;
      report("connecting");
      tokenClient.requestAccessToken({ prompt: "" });
    }

    function enqueue({ upserts = [], deletes = [] }) {
      queue(upserts, deletes);
      if (token && navigator.onLine) sync();
    }

    function disconnect() {
      if (token && window.google?.accounts?.oauth2) {
        google.accounts.oauth2.revoke(token, () => {});
      }
      token = "";
      expired = false;
      report("local", "已断开 Google，同步资料仍保留在 Sheet");
    }

    function openSheet() {
      if (meta.sheetId) {
        window.open(
          `https://docs.google.com/spreadsheets/d/${meta.sheetId}/edit`,
          "_blank",
          "noopener",
        );
      }
    }

    function init() {
      saveMeta();
      report(config.googleClientId ? "local" : "unconfigured");
      window.addEventListener("online", () => sync());
      window.addEventListener("offline", () =>
        report("offline", "离线中，活动已保存在此设备"),
      );
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") sync();
      });
    }

    return {
      init,
      connect,
      sync,
      enqueue,
      disconnect,
      openSheet,
      isConnected: () => Boolean(token),
    };
  }

  window.GoogleTimetableSync = { create, __test: { reduceRows, revisionFor } };
})();
