/* Advice Call side panel
 * - Runs inside a Teams meeting side panel (or in a normal browser for testing).
 * - Finds the nrq_advicecall record linked to the meeting (via the meeting chat thread id).
 * - Shows the fields and saves changes to Dataverse as the user types.
 *
 * Test outside Teams: open index.html?id=<advice call record GUID> in a browser.
 */
(function () {
  "use strict";

  var S = window.APP_SETTINGS;
  var DV = S.dataverseUrl.replace(/\/+$/, "");
  var API = DV + "/api/data/v9.2";
  var SCOPES = [DV + "/user_impersonation"];
  var TABLE = "nrq_advicecalls";
  var ID_FIELD = "nrq_advicecallid";

  // Choice values from the DemoAdviceCall solution
  var CATEGORY = [
    [875810000, "Commercial advice"],
    [875810001, "Legal Advice"],
    [875810002, "Complaint"],
    [875810003, "Other"]
  ];
  var SATISFACTION = [
    [875810000, "Bad"],
    [875810001, "Neutral"],
    [875810002, "Good"]
  ];

  var SELECT = [
    "nrq_advicecallid", "nrq_name", "nrq_startdate", "nrq_enddate",
    "nrq_category", "nrq_satisfactorylevel", "nrq_problem",
    "nrq_guidancereference", "nrq_nextactionpoints", "_nrq_contactid_value"
  ].join(",");

  // Sample data for design preview: open index.html?preview=1 (add &theme=dark to see dark mode)
  function sampleRecord() {
    var start = new Date();
    start.setDate(start.getDate() + 1);
    start.setHours(10, 0, 0, 0);
    var end = new Date(start.getTime() + 60 * 60 * 1000);
    return {
      nrq_advicecallid: "00000000-0000-0000-0000-000000000000",
      nrq_name: "Alex Fornier / Commercial advice",
      nrq_startdate: start.toISOString(),
      nrq_enddate: end.toISOString(),
      nrq_category: 875810000,
      nrq_satisfactorylevel: 875810002,
      nrq_problem: "Customer wants to merge two price lists and keep the customer-specific discounts. Unsure which items are affected.",
      nrq_guidancereference: "Pricing guide, section 4.2",
      nrq_nextactionpoints: "Send overview of affected items\nPlan follow-up call next week",
      "_nrq_contactid_value@OData.Community.Display.V1.FormattedValue": "Alex Fornier"
    };
  }

  var state = {
    preview: false,
    inTeams: false,
    ctx: null,
    pca: null,
    record: null,
    pending: {},
    timer: null,
    saving: Promise.resolve(),
    diag: {}
  };

  function $(id) { return document.getElementById(id); }

  function showView(name) {
    ["loading", "signin", "notfound", "error", "form"].forEach(function (v) {
      $("view-" + v).hidden = (v !== name);
    });
  }

  function showError(message) {
    $("error-text").textContent = message;
    showView("error");
  }

  function updateDiag() {
    $("diag-text").textContent = JSON.stringify(state.diag, null, 2);
  }

  function setStatus(text, kind) {
    var el = $("status");
    el.textContent = text;
    el.className = "status " + (kind || "muted");
  }

  function withTimeout(promise, ms) {
    return Promise.race([
      promise,
      new Promise(function (_, reject) {
        setTimeout(function () { reject(new Error("timeout")); }, ms);
      })
    ]);
  }

  function applyTheme(theme) {
    document.body.classList.remove("theme-dark", "theme-contrast");
    if (theme === "dark") document.body.classList.add("theme-dark");
    if (theme === "contrast") document.body.classList.add("theme-contrast");
  }

  // ---------- Teams + MSAL setup ----------

  async function initTeams() {
    if (window.self === window.top) return false; // not embedded, so not Teams
    try {
      await withTimeout(microsoftTeams.app.initialize(), 5000);
      state.ctx = await microsoftTeams.app.getContext();
      state.inTeams = true;
      if (state.ctx.app && state.ctx.app.theme) applyTheme(state.ctx.app.theme);
      microsoftTeams.app.registerOnThemeChangeHandler(applyTheme);
      microsoftTeams.app.notifyAppLoaded();
      microsoftTeams.app.notifySuccess();
      return true;
    } catch (e) {
      state.diag.teamsInitError = String(e && e.message || e);
      return false;
    }
  }

  async function initMsal() {
    var config = {
      auth: {
        clientId: S.clientId,
        authority: "https://login.microsoftonline.com/" + S.tenantId,
        redirectUri: S.siteUrl,
        supportsNestedAppAuth: true
      },
      cache: { cacheLocation: "sessionStorage" }
    };
    if (state.inTeams) {
      state.pca = await msal.createNestablePublicClientApplication(config);
    } else {
      state.pca = await msal.createStandardPublicClientApplication(config);
    }
  }

  function loginHint() {
    var u = state.ctx && state.ctx.user;
    return u ? (u.loginHint || u.userPrincipalName || undefined) : undefined;
  }

  // Returns an access token for Dataverse. If interactive is false and the user
  // must sign in, throws an error with needsSignIn = true.
  async function getToken(interactive, forceRefresh) {
    var pca = state.pca;
    var account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    var result;
    if (!account && !interactive && !loginHint()) {
      // Test mode in a normal browser: nothing to sign in silently with.
      var needs = new Error("Sign-in required");
      needs.needsSignIn = true;
      throw needs;
    }
    try {
      if (account) {
        result = await pca.acquireTokenSilent({ scopes: SCOPES, account: account, forceRefresh: !!forceRefresh });
      } else {
        result = await pca.ssoSilent({ scopes: SCOPES, loginHint: loginHint() });
      }
    } catch (silentError) {
      if (!interactive) {
        var err = new Error("Sign-in required");
        err.needsSignIn = true;
        err.inner = silentError;
        throw err;
      }
      result = await pca.acquireTokenPopup({ scopes: SCOPES, loginHint: loginHint() });
    }
    pca.setActiveAccount(result.account);
    state.diag.signedInAs = result.account && result.account.username;
    return result.accessToken;
  }

  // ---------- Dataverse calls ----------

  async function dv(method, path, body, isRetry) {
    var token = await getToken(false, !!isRetry);
    var headers = {
      "Authorization": "Bearer " + token,
      "Accept": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0"
    };
    if (method === "GET") {
      headers["Prefer"] = 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"';
    } else {
      headers["Content-Type"] = "application/json; charset=utf-8";
    }
    if (method === "PATCH") headers["If-Match"] = "*"; // update only, never create

    var res = await fetch(API + path, {
      method: method,
      headers: headers,
      body: body ? JSON.stringify(body) : undefined
    });

    if (res.status === 401 && !isRetry) return dv(method, path, body, true);
    if (!res.ok) {
      var text = await res.text();
      var msg = text;
      try { msg = JSON.parse(text).error.message; } catch (e) { /* keep raw text */ }
      throw new Error(res.status + " " + (msg || res.statusText));
    }
    return res.status === 204 ? null : res.json();
  }

  function odataString(value) {
    return "'" + String(value).replace(/'/g, "''") + "'";
  }

  async function findRecord() {
    var params = new URLSearchParams(window.location.search);
    var id = params.get("id");
    if (id) {
      state.diag.lookup = "by record id (test mode)";
      return dv("GET", "/" + TABLE + "(" + encodeURIComponent(id) + ")?$select=" + SELECT);
    }

    var chatId = state.ctx && state.ctx.chat && state.ctx.chat.id;
    state.diag.chatId = chatId || null;
    state.diag.meetingId = (state.ctx && state.ctx.meeting && state.ctx.meeting.id) || null;
    state.diag.frameContext = (state.ctx && state.ctx.page && state.ctx.page.frameContext) || null;
    if (!chatId) return null;

    state.diag.lookup = "by chat thread id";
    var filter = "nrq_chatthreadid eq " + odataString(chatId);
    var data = await dv("GET", "/" + TABLE + "?$select=" + SELECT +
      "&$filter=" + encodeURIComponent(filter) + "&$top=1");
    return data && data.value && data.value.length ? data.value[0] : null;
  }

  // ---------- Form ----------

  function fmtDate(iso) {
    if (!iso) return "";
    try {
      return new Date(iso).toLocaleString([], { dateStyle: "medium", timeStyle: "short" });
    } catch (e) {
      return iso;
    }
  }

  function renderForm(rec) {
    $("f-name").textContent = rec.nrq_name || "Advice call";
    var contact = rec["_nrq_contactid_value@OData.Community.Display.V1.FormattedValue"];
    $("f-contact").textContent = contact ? "Contact: " + contact : "";
    var when = fmtDate(rec.nrq_startdate);
    var until = rec.nrq_enddate ? " to " + new Date(rec.nrq_enddate).toLocaleTimeString([], { timeStyle: "short" }) : "";
    $("f-when").textContent = when ? when + until : "";

    // Category
    var cat = $("f-category");
    cat.innerHTML = "";
    var empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "Select...";
    cat.appendChild(empty);
    CATEGORY.forEach(function (c) {
      var o = document.createElement("option");
      o.value = String(c[0]);
      o.textContent = c[1];
      cat.appendChild(o);
    });
    cat.value = rec.nrq_category == null ? "" : String(rec.nrq_category);
    cat.onchange = function () {
      queueSave("nrq_category", cat.value === "" ? null : parseInt(cat.value, 10), 100);
    };

    // Satisfactory level (segmented radio buttons)
    var sat = $("f-satisfaction");
    sat.innerHTML = "";
    SATISFACTION.forEach(function (s) {
      var label = document.createElement("label");
      var input = document.createElement("input");
      input.type = "radio";
      input.name = "satisfaction";
      input.value = String(s[0]);
      input.checked = rec.nrq_satisfactorylevel === s[0];
      input.onchange = function () {
        queueSave("nrq_satisfactorylevel", s[0], 100);
      };
      var span = document.createElement("span");
      span.textContent = s[1];
      label.appendChild(input);
      label.appendChild(span);
      sat.appendChild(label);
    });

    // Text areas
    bindText("f-problem", "nrq_problem", rec);
    bindText("f-guidance", "nrq_guidancereference", rec);
    bindText("f-next", "nrq_nextactionpoints", rec);

    setStatus("Changes are saved automatically.", "muted");
    showView("form");
  }

  function bindText(elementId, field, rec) {
    var el = $(elementId);
    el.value = rec[field] || "";
    el.oninput = function () { queueSave(field, el.value, 900); };
    el.onblur = function () { flush(); };
  }

  // ---------- Saving ----------

  function queueSave(field, value, delay) {
    state.pending[field] = value;
    setStatus("Unsaved changes...", "muted");
    clearTimeout(state.timer);
    state.timer = setTimeout(flush, delay);
  }

  function flush() {
    clearTimeout(state.timer);
    state.saving = state.saving.then(doSave, doSave);
    return state.saving;
  }

  async function doSave() {
    if (state.preview) {
      state.pending = {};
      setStatus("Preview mode: changes are not saved.", "muted");
      return;
    }
    var keys = Object.keys(state.pending);
    if (!keys.length || !state.record) return;
    var body = {};
    keys.forEach(function (k) { body[k] = state.pending[k]; });
    setStatus("Saving...", "muted");
    try {
      await dv("PATCH", "/" + TABLE + "(" + state.record[ID_FIELD] + ")", body);
      keys.forEach(function (k) {
        if (state.pending[k] === body[k]) delete state.pending[k];
      });
      if (Object.keys(state.pending).length) {
        flush();
      } else {
        setStatus("Saved " + new Date().toLocaleTimeString([], { timeStyle: "short" }), "ok");
      }
    } catch (e) {
      if (e.needsSignIn) {
        setStatus("Not saved: please sign in again.", "err");
        showView("signin");
      } else {
        setStatus("Not saved: " + e.message + " (will retry on your next change)", "err");
      }
    }
  }

  // ---------- Start ----------

  async function load() {
    showView("loading");
    try {
      await getToken(false);
    } catch (e) {
      if (e.needsSignIn) {
        $("signin-note").textContent = state.inTeams ? "" : "Test mode: a sign-in window will open.";
        showView("signin");
        return;
      }
      showError(e.message);
      return;
    }
    try {
      var rec = await findRecord();
      updateDiag();
      if (!rec) {
        showView("notfound");
        return;
      }
      state.record = rec;
      state.diag.recordId = rec[ID_FIELD];
      updateDiag();
      renderForm(rec);
    } catch (e) {
      updateDiag();
      showError(e.message);
    }
  }

  async function init() {
    // If this page was opened as the sign-in popup return, let MSAL finish and close it.
    var returningFromPopup = window.opener && window.opener !== window &&
      /[#&](code|error)=/.test(window.location.hash || "");
    if (returningFromPopup) {
      await msal.createStandardPublicClientApplication({
        auth: { clientId: S.clientId, authority: "https://login.microsoftonline.com/" + S.tenantId, redirectUri: S.siteUrl },
        cache: { cacheLocation: "sessionStorage" }
      });
      return;
    }

    var urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get("theme")) applyTheme(urlParams.get("theme"));

    if (urlParams.get("preview")) {
      state.preview = true;
      state.diag.panelVersion = S.version;
      state.diag.mode = "design preview (sample data)";
      updateDiag();
      renderForm(sampleRecord());
      setStatus("Preview mode: changes are not saved.", "muted");
      return;
    }

    state.diag.panelVersion = S.version;
    await initTeams();
    state.diag.inTeams = state.inTeams;
    if (state.ctx && state.ctx.user) state.diag.user = state.ctx.user.userPrincipalName || null;

    try {
      await initMsal();
    } catch (e) {
      updateDiag();
      showError("Could not start sign-in: " + (e && e.message || e));
      return;
    }

    $("btn-signin").onclick = async function () {
      try {
        await getToken(true);
        await load();
      } catch (e) {
        showError("Sign-in failed: " + (e && e.message || e));
      }
    };
    $("btn-retry").onclick = load;

    window.addEventListener("pagehide", function () { flush(); });
    document.addEventListener("visibilitychange", function () {
      if (document.visibilityState === "hidden") flush();
    });

    updateDiag();
    await load();
  }

  init().catch(function (e) {
    showError(String(e && e.message || e));
  });
})();
