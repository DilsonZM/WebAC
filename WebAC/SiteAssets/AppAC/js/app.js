const app = (() => {
  let gerenciaChart = null;
  let criticidadChart = null;
  let advCoverageChart = null;
  let advCriticidadChart = null;
  let advAltaChart = null;
  let advTrendChart = null;
  const CONFIG = {
    listNames: {
      equipos: "EquiposAC",
      escenarios: "EscenariosAC",
      evaluaciones: "EvaluacionesAC",
      roles: "RolesAC",
      editRequests: "SolicitudesEdicionAC"
    },
    pageSize: 2000,
    adminGroups: ["AC-Admins", "SGIA-Admins"],
    preguntas: ["FO", "FIN", "HSE", "MA", "SOC", "DDHH", "REP", "LEGAL", "TI", "FF"],
    preguntaLabels: {
      FO: "FLEXIBILIDAD OPERACIONAL",
      FIN: "FINANCIERO",
      HSE: "SALUD Y SEGURIDAD",
      MA: "MEDIO AMBIENTE",
      SOC: "SOCIAL Y CULTURAL",
      DDHH: "DERECHOS HUMANOS",
      REP: "REPUTACIÓN E IMAGEN",
      LEGAL: "LEGAL",
      TI: "TI",
      FF: "FRECUENCIA DE FALLA"
    },
    escala: [1, 2, 3, 4, 5],
    versionMetodologia: "1.0",
    equiposJsonMax: 2000,
    maxEscenarios: 200,
    requireScenarioSearch: false
  };

  const goalPct = 90;

  const state = {
    user: null,
    isAdmin: false,
    evalFields: null,
    adminEval: { page: 0, pageSize: 5, items: [] },
    dashEval: { page: 0, pageSize: 5, items: [] },
    misEval: { page: 0, pageSize: 5, items: [], filtered: [], updatingOptions: false },
    allEval: { page: 0, pageSize: 8, items: [] },
    evalDetalle: { page: 0, pageSize: 10, items: [], filtered: [] },
    role: { canEdit: false, canAdmin: false },
    lastView: null,
    currentEvalId: null,
    currentEvalItem: null,
    escFields: null,
    evalJsonIndex: { loadedAt: 0, items: [], byEquip: new Map(), equipSet: new Set() },
    equipos: [],
    equiposMap: new Map(),
    distinct: { fleet: [], proceso: [], egi: [], gerencia: [], superintendencia: [], unidadProceso: [] },
    wizard: {
      nivel: "",
      fleet: "",
      proceso: "",
      egi: "",
      equipo: null,
      escenarioCandidates: [],
      equiposCubiertos: 0,
      escenario: null,
      escenarioSearch: "",
      escenarioThrottled: false,
      escenarioPage: 0,
      escenarioPageSize: 5,
      respuestas: {},
      justificacion: "",
      existingEvalId: null,
      coverage: null,
      eventoAnalizado: ""
    },
    search: { term: "", page: 0, pageSize: 5, results: [] },
    adminEditingId: null,
    dashboardSummary: null,
    dashboardAdvanced: { gerencia: "", superintendencia: "", unidad: "", fleet: "", proceso: "", egi: "", dim: "GERENCIA" }
  };

  const el = (id) => document.getElementById(id);
  const fmt = (value) => value || "-";
  const normalizeEquip = (value) => String(value || "").trim().toUpperCase();
  const normalizeDim = (value) => String(value || "").trim().toUpperCase();
  const escapeOData = (value) => String(value || "").replace(/'/g, "''");
    function animateNumber(elm, to, { duration = 600, formatter } = {}) {
      if (!elm || !Number.isFinite(to)) return;
      const fromText = String(elm.textContent || "").replace(/[^0-9.\-]/g, "");
      const from = Number(fromText);
      const start = performance.now();
      const fmtLocal = formatter || ((v) => String(Math.round(v)));
      const fromVal = Number.isFinite(from) ? from : 0;
      function tick(now) {
        const t = Math.min(1, (now - start) / duration);
        const val = fromVal + (to - fromVal) * t;
        elm.textContent = fmtLocal(val);
        if (t < 1) requestAnimationFrame(tick);
      }
      requestAnimationFrame(tick);
    }

    function setAnimated(id, value, formatter) {
      const node = el(id);
      if (!node) return;
      animateNumber(node, value, { formatter });
    }
  const stopwords = new Set(["de","la","que","el","en","y","a","los","del","se","las","por","un","para","con","no","una","su","al","lo","como","más","pero","sus","le","ya","o","este","sí","porque","esta","entre","cuando","muy","sin","sobre","también","me","hasta","hay","donde","quien","desde","todo","nos","durante","todos","uno","les","ni","contra","otros","ese","eso","ante","ellos","e","esto","mí","antes","algunos","qué","unos","yo","otro","otras","otra","él","tanto","esa","estos","mucho","quienes","nada","muchos","cual","poco","ella","estar","estas","algunas","algo","nosotros","mi","mis","tú","te","ti","tu","tus","ellas","nosotras","vosostros","vosostras","os","mío","mía","míos","mías","tuyo","tuya","tuyos","tuyas","suyo","suya","suyos","suyas","nuestro","nuestra","nuestros","nuestras","vuestro","vuestra","vuestros","vuestras","esos","esas","estoy","está","estamos","están","esté","sea","son","fue","han","ser","tener","hace"]);

  function getInitials(name) {
    const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return "AC";
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function setUserProfile(user) {
    const nameEl = el("user-name");
    const emailEl = el("user-email");
    const avatarEl = el("user-avatar");
    const initialsEl = el("user-initials");
    const photoEl = el("user-photo");
    if (!nameEl || !emailEl || !avatarEl || !initialsEl || !photoEl) return;
    const displayName = user?.Title || "Usuario";
    const email = user?.Email || "";
    nameEl.textContent = displayName;
    emailEl.textContent = email || "-";
    initialsEl.textContent = getInitials(displayName);
    if (email) {
      const webUrl = spHttp.getWebUrl();
      const photoUrl = `${webUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
      photoEl.onload = () => avatarEl.classList.add("has-photo");
      photoEl.onerror = () => avatarEl.classList.remove("has-photo");
      photoEl.src = photoUrl;
    }
  }

  function resolveFieldInternal(fieldsMap, ...candidates) {
    const keys = Object.keys(fieldsMap || {});
    const lower = new Map(keys.map(k => [k.toLowerCase(), k]));
    for (const name of candidates) {
      const direct = fieldsMap[name];
      if (direct) return name;
      const key = lower.get(String(name).toLowerCase());
      if (key) return key;
    }
    return null;
  }

  function resolveQuestionFields(fieldsMap) {
    const map = {};
    const keys = Object.keys(fieldsMap || {});
    const normalize = (v) => String(v || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/\p{Diacritic}/gu, "")
      .replace(/[^a-z0-9]/g, "");
    const normKeyMap = new Map(keys.map(k => [normalize(k), k]));

    CONFIG.preguntas.forEach(k => {
      let found = resolveFieldInternal(fieldsMap, k);
      if (!found) {
        const label = CONFIG.preguntaLabels[k] || "";
        const byKey = normKeyMap.get(normalize(k));
        const byLabel = normKeyMap.get(normalize(label));
        found = byKey || byLabel || null;
      }
      map[k] = found;
    });
    return map;
  }

  async function getEvaluacionesFieldsDetailed() {
    const webUrl = spHttp.getWebUrl();
    const safeName = CONFIG.listNames.evaluaciones.replace(/'/g, "''");
    const url = `${webUrl}/_api/web/lists/GetByTitle('${safeName}')/fields?$select=InternalName,Title,TypeAsString,Hidden`;
    try {
      const data = await spHttp.get(url);
      return (data.value || []).filter(f => !f.Hidden);
    } catch {
      return [];
    }
  }

  function resolveQuestionFieldsFromDetails(fields) {
    const map = {};
    const normalize = (v) => String(v || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/\p{Diacritic}/gu, "")
      .replace(/[^a-z0-9]/g, "");
    const byNormTitle = new Map(fields.map(f => [normalize(f.Title), f.InternalName]));
    const byNormInternal = new Map(fields.map(f => [normalize(f.InternalName), f.InternalName]));

    CONFIG.preguntas.forEach(k => {
      const label = CONFIG.preguntaLabels[k] || "";
      let found = byNormTitle.get(normalize(k)) || byNormInternal.get(normalize(k))
        || byNormTitle.get(normalize(label)) || byNormInternal.get(normalize(label)) || null;
      if (!found && k === "FF") {
        const ffField = fields.find(f => {
          const t = normalize(f.Title);
          const n = normalize(f.InternalName);
          return (t.includes("frecuencia") && t.includes("falla")) || (n.includes("frecuencia") && n.includes("falla"));
        });
        if (ffField) found = ffField.InternalName;
      }
      map[k] = found;
    });
    return map;
  }

  function getQuestionValue(item, key) {
    if (!item) return null;
    const map = item._qMap || {};
    const internal = map[key] || key;
    if (item[internal] !== undefined && item[internal] !== null && item[internal] !== "") return item[internal];
    if (item[key] !== undefined && item[key] !== null && item[key] !== "") return item[key];
    if (key === "FF") {
      if (item._ffValue !== undefined) return item._ffValue;
      const maybe = Object.keys(item).find(k => /frecuencia|falla/i.test(k));
      if (maybe) return item[maybe];
    }
    return null;
  }

  async function getFieldValuesAsText(listName, id) {
    const url = `${spHttp.listItemsUrl(listName)}(${id})?$select=FieldValuesAsText&$expand=FieldValuesAsText`;
    const data = await spHttp.get(url);
    return data && data.FieldValuesAsText ? data.FieldValuesAsText : null;
  }

  function swalInfo(title, text) {
    if (window.Swal) return Swal.fire({ icon: "info", title, text, confirmButtonText: "Entendido" });
    alert(text || title);
  }

  function swalError(title, text) {
    if (window.Swal) return Swal.fire({ icon: "error", title, text, confirmButtonText: "Entendido" });
    alert(text || title);
  }

  function swalSuccess(title, text) {
    if (window.Swal) return Swal.fire({ icon: "success", title, text, confirmButtonText: "OK" });
    alert(text || title);
  }

  async function withMetadata(listName, payload) {
    if (payload && payload.__metadata) return payload;
    const typeName = await spHttp.getListItemEntityType(listName);
    return { __metadata: { type: typeName }, ...payload };
  }

  function extractPreview(text, maxSentences = 2) {
    if (!text) return "";
    const sentences = text
      .replace(/\s+/g, " ")
      .split(/(?<=[.!?])\s+/)
      .filter(s => s && s.length > 20);
    if (sentences.length <= maxSentences) return sentences.join(" ");

    const wordFreq = new Map();
    const tokenize = (s) => s.toLowerCase().replace(/[^a-záéíóúñ0-9\s]/gi, " ").split(/\s+/).filter(w => w.length > 2 && !stopwords.has(w));
    sentences.forEach(s => {
      tokenize(s).forEach(w => wordFreq.set(w, (wordFreq.get(w) || 0) + 1));
    });

    const scored = sentences.map((s, idx) => {
      const score = tokenize(s).reduce((sum, w) => sum + (wordFreq.get(w) || 0), 0);
      return { s, idx, score };
    });

    const top = scored.sort((a, b) => b.score - a.score).slice(0, maxSentences).sort((a, b) => a.idx - b.idx);
    return top.map(t => t.s).join(" ");
  }

  function isAdmin(groups) {
    if (!groups) return false;
    const names = groups.map(g => g.Title.toLowerCase());
    return CONFIG.adminGroups.some(g => names.includes(g.toLowerCase()));
  }

  async function loadEquipos() {
    const data = await catalogService.loadEquipos(CONFIG.pageSize);
    state.equipos = data.equipos;
    state.equiposMap = data.equiposMap;
    state.distinct = data.distinct;
    fillSelect(el("sel-fleet"), state.distinct.fleet, "Seleccione...");
    fillSelect(el("sel-proceso"), state.distinct.proceso, "Seleccione...");
    fillSelect(el("sel-egi"), state.distinct.egi, "Seleccione...");
    fillSelect(el("mis-fleet"), state.distinct.fleet, "Todos");
    fillSelect(el("mis-proceso"), state.distinct.proceso, "Todos");
    fillSelect(el("mis-egi"), state.distinct.egi, "Todos");
    fillSelect(el("all-fleet"), state.distinct.fleet, "Todos");
    fillSelect(el("all-proceso"), state.distinct.proceso, "Todos");
    fillSelect(el("all-egi"), state.distinct.egi, "Todos");
    fillSelect(el("admin-fleet"), state.distinct.fleet, "Todos");
    fillSelect(el("admin-proceso"), state.distinct.proceso, "Todos");
    fillSelect(el("admin-egi"), state.distinct.egi, "Todos");
  }

  function fillSelect(select, options, placeholder) {
    if (!select) return;
    select.innerHTML = "";
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = placeholder || "Seleccione...";
    select.appendChild(opt);
    options.forEach(value => {
      const o = document.createElement("option");
      o.value = value;
      o.textContent = value;
      select.appendChild(o);
    });
  }

  function setView(viewId) {
    document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
    const view = el(viewId);
    if (view) view.classList.add("active");
    document.querySelectorAll(".side-btn[data-view]").forEach(btn => {
      btn.classList.toggle("active", btn.dataset.view === viewId);
    });
    const titleEl = el("page-title");
    const subtitleEl = el("page-subtitle");
    if (titleEl && subtitleEl) {
      const map = {
        "view-dashboard": { title: "Dashboard", subtitle: "Resumen general" },
        "view-dashboard-advanced": { title: "Inicio", subtitle: "Resumen general" },
        "view-wizard": { title: "Analisis de criticidad", subtitle: "Nuevo análisis" },
        "view-mis": { title: "Mis análisis", subtitle: "Historial personal" },
        "view-all": { title: "Todos los análisis", subtitle: "Historial general" },
        "view-admin": { title: "Administración", subtitle: "Solicitudes y análisis" },
        "view-evaluacion": { title: "Análisis", subtitle: "Detalle de análisis" }
      };
      const meta = map[viewId] || { title: "Análisis", subtitle: "" };
      titleEl.textContent = meta.title;
      subtitleEl.textContent = meta.subtitle;
    }
  }

  function setLoading(isLoading) {
    const overlay = el("loading-overlay");
    if (overlay) overlay.classList.toggle("active", isLoading);
    document.body.classList.toggle("loading", isLoading);
  }

  function setStep(step) {
    const steps = ["step-1", "step-2", "step-3", "step-4", "step-5", "step-6"];
    steps.forEach((id, idx) => {
      el(id).classList.toggle("hidden", idx + 1 !== step);
    });
    document.querySelectorAll("#wizard-steps .step").forEach((s, idx) => {
      s.classList.toggle("active", idx + 1 === step);
    });
  }

  const getDimValue = (equip, dim) => {
    if (!equip) return "";
    if (dim === "GERENCIA") return equip.BranchGerencia || "";
    if (dim === "SUPERINTENDENCIA") return equip.SiteSuperintendencia || "";
    if (dim === "UNIDADPROCESO") return equip.UnidadProceso || "";
    if (dim === "FLEET") return equip.Fleet || "";
    if (dim === "PROCESO") return equip.ProcesoSistema || "";
    if (dim === "EGI") return equip.EGI || "";
    return "";
  };

  const resolveEquipMetaFromItem = (item) => {
    const cov = parseJsonList(item?.EquiposCubiertosJson);
    const first = cov.find(Boolean);
    const fromCov = normalizeEquip(first?.EquipNo || first?.Title || first || "");
    const fromEquipNo = normalizeEquip(item?.EquipNo || "");
    const fromTarget = normalizeEquip(item?.TargetValor || "");
    const key = fromCov || fromEquipNo || fromTarget;
    let equip = key ? state.equiposMap.get(key) : null;
    if (!equip && item?.TargetTipo && item?.TargetValor) {
      const tipo = String(item.TargetTipo || "").toUpperCase();
      const valor = String(item.TargetValor || "").trim().toUpperCase();
      if (tipo === "FLEET") {
        equip = state.equipos.find(e => String(e.Fleet || "").trim().toUpperCase() === valor) || null;
      } else if (tipo === "PROCESO_SISTEMA" || tipo === "PROCESOSISTEMA") {
        equip = state.equipos.find(e => String(e.ProcesoSistema || "").trim().toUpperCase() === valor) || null;
      } else if (tipo === "EGI") {
        equip = state.equipos.find(e => String(e.EGI || "").trim().toUpperCase() === valor) || null;
      } else if (tipo === "EQUIPO") {
        equip = state.equipos.find(e => normalizeEquip(e.EquipNo || e.Title) === normalizeEquip(valor)) || null;
      }
    }
    return {
      gerencia: equip?.BranchGerencia || item?.Gerencia || "-",
      superintendencia: equip?.SiteSuperintendencia || item?.Superintendencia || "-",
      unidadProceso: equip?.UnidadProceso || equip?.ProcesoSistema || item?.UnidadProceso || item?.ProcesoSistema || "-"
    };
  };

  function updateAdvancedKpis(summary) {
    if (!summary) return;
    const goalPct = 90;
    const {
      totalFleets, totalProcesos, totalEgi, totalEquipos,
      coveredFleets, coveredProcesos, coveredEgi, coveredEquipos,
      pctFleets, pctProcesos, pctEgi, pctEquipos,
      pendingFleets, pendingProcesos, pendingEgi, pendingEquipos
    } = summary;
    const setText = (id, value) => { const node = el(id); if (node) node.textContent = value; };
    const setAnimatedLocal = (id, value, formatter) => {
      const node = el(id); if (node) setAnimated(id, value, formatter);
    };
    setAnimatedLocal("adv-kpi-flotas-ok", coveredFleets || 0, v => `${Math.round(v)}`);
    setAnimatedLocal("adv-kpi-equipos-ok", coveredEquipos || 0, v => `${Math.round(v)}`);
    setAnimatedLocal("adv-kpi-procesos-ok", coveredProcesos || 0, v => `${Math.round(v)}`);
    setAnimatedLocal("adv-kpi-egi-ok", coveredEgi || 0, v => `${Math.round(v)}`);
    setText("adv-kpi-flotas-total", `de ${totalFleets}`);
    setText("adv-kpi-equipos-total", `de ${totalEquipos}`);
    setText("adv-kpi-procesos-total", `de ${totalProcesos}`);
    setText("adv-kpi-egi-total", `de ${totalEgi}`);
    setAnimatedLocal("adv-kpi-flotas-pct", pctFleets, v => `▲ ${v.toFixed(2)}%`);
    setAnimatedLocal("adv-kpi-equipos-pct", pctEquipos, v => `▲ ${v.toFixed(2)}%`);
    setAnimatedLocal("adv-kpi-procesos-pct", pctProcesos, v => `▲ ${v.toFixed(2)}%`);
    setAnimatedLocal("adv-kpi-egi-pct", pctEgi, v => `▲ ${v.toFixed(2)}%`);
    setAnimatedLocal("adv-kpi-flotas-pend", pendingFleets, v => `▼ ${Math.round(v)}`);
    setAnimatedLocal("adv-kpi-equipos-pend", pendingEquipos, v => `▼ ${Math.round(v)}`);
    setAnimatedLocal("adv-kpi-procesos-pend", pendingProcesos, v => `▼ ${Math.round(v)}`);
    setAnimatedLocal("adv-kpi-egi-pend", pendingEgi, v => `▼ ${Math.round(v)}`);

    const setCardDanger = (cardId, pct) => {
      const card = el(cardId);
      if (!card) return;
      card.classList.toggle("metric-danger", pct < goalPct);
    };
    const setRowDanger = (pctId, pct) => {
      const node = el(pctId);
      if (!node) return;
      const row = node.closest(".metric-row");
      if (!row) return;
      row.classList.toggle("danger", pct < goalPct);
    };
    setCardDanger("adv-card-flotas", pctFleets);
    setCardDanger("adv-card-equipos", pctEquipos);
    setCardDanger("adv-card-procesos", pctProcesos);
    setCardDanger("adv-card-egis", pctEgi);
    setRowDanger("adv-kpi-flotas-pct", pctFleets);
    setRowDanger("adv-kpi-equipos-pct", pctEquipos);
    setRowDanger("adv-kpi-procesos-pct", pctProcesos);
    setRowDanger("adv-kpi-egi-pct", pctEgi);
  }

  function showError(message) {
    const box = el("wizard-error");
    box.textContent = message;
    box.classList.toggle("hidden", !message);
  }

  function resetWizard() {
    state.wizard = { nivel: "", fleet: "", proceso: "", egi: "", equipo: null, escenarioCandidates: [], equiposCubiertos: 0, escenario: null, escenarioSearch: "", escenarioThrottled: false, escenarioPage: 0, escenarioPageSize: 5, respuestas: {}, justificacion: "", existingEvalId: null, coverage: null, eventoAnalizado: "" };
    el("nivel-select").value = "";
    el("sel-fleet").value = "";
    el("sel-proceso").value = "";
    el("sel-egi").value = "";
    el("txt-equipo").value = "";
    el("equipo-results").innerHTML = "";
    el("evento-analizado").value = "";
    if (el("evento-analizado-edit")) el("evento-analizado-edit").value = "";
    el("escenario-list").innerHTML = "";
    el("escenario-meta").textContent = "-";
    el("equipo-count-target").textContent = "-";
    el("escenario-search").value = "";
    el("escenario-count").textContent = "";
    el("justificacion").value = "";
    el("to-step-2").disabled = true;
    el("to-step-3").disabled = true;
    el("to-step-4").disabled = true;
    showError("");
    setStep(1);
    clearExistingEvaluacionNotice();
  }

  function resetFilters() {
    state.wizard.fleet = "";
    state.wizard.proceso = "";
    state.wizard.egi = "";
    state.wizard.equipo = null;
    state.wizard.coverage = null;
    state.wizard.existingEvalId = null;
    el("sel-fleet").value = "";
    el("sel-proceso").value = "";
    el("sel-egi").value = "";
    el("txt-equipo").value = "";
    state.search.page = 0;
    state.search.results = [];
    el("equipo-results").innerHTML = "";
    el("equipo-count").textContent = "";
    el("equipo-page").value = "1";
    el("equipo-page-total").textContent = "";
    el("to-step-3").disabled = true;
    clearExistingEvaluacionNotice();
  }

  function buildPreguntas() {
    const container = el("preguntas");
    if (!container) return;
    container.innerHTML = "";
    CONFIG.preguntas.forEach(key => {
      const wrapper = document.createElement("div");
      const label = document.createElement("label");
      label.className = "question-label";
      label.innerHTML = `<span class="q-code">${key}</span><span class="q-text">${CONFIG.preguntaLabels[key] || ""}</span>`;
      const select = document.createElement("select");
      select.className = "form-select";
      select.dataset.key = key;
      select.innerHTML = `<option value="">Seleccione</option>` + CONFIG.escala.map(v => `<option value="${v}">${v}</option>`).join("");
      wrapper.appendChild(label);
      wrapper.appendChild(select);
      container.appendChild(wrapper);
    });
  }

  function buildEvalEditPreguntas() {
    const container = el("eval-edit-preguntas");
    if (!container) return;
    container.innerHTML = "";
    CONFIG.preguntas.forEach(key => {
      const wrapper = document.createElement("div");
      const label = document.createElement("label");
      label.className = "question-label";
      label.innerHTML = `<span class="q-code">${key}</span><span class="q-text">${CONFIG.preguntaLabels[key] || ""}</span>`;
      const select = document.createElement("select");
      select.className = "form-select";
      select.dataset.key = key;
      select.innerHTML = `<option value="">Seleccione</option>` + CONFIG.escala.map(v => `<option value="${v}">${v}</option>`).join("");
      wrapper.appendChild(label);
      wrapper.appendChild(select);
      container.appendChild(wrapper);
    });
  }

  async function loadUserRoles() {
    try {
      const fieldsMap = await spHttp.getListFields(CONFIG.listNames.roles);
      const userField = resolveFieldInternal(fieldsMap, "Usuario", "User", "Persona");
      if (!userField || !state.user) return;
      const filter = `${userField}Id eq ${state.user.Id}`;
      const selectParts = ["Id", `${userField}/Id`, `${userField}/Title`];
      if (fieldsMap.PuedeEditar) selectParts.push("PuedeEditar");
      if (fieldsMap.PuedeAdministrar) selectParts.push("PuedeAdministrar");
      const url = spHttp.listItemsUrl(CONFIG.listNames.roles, `$select=${selectParts.join(",")}&$expand=${userField}&$filter=${encodeURIComponent(filter)}&$top=1`);
      const data = await spHttp.get(url);
      const item = (data.value || [])[0];
      state.role.canEdit = !!(item && item.PuedeEditar);
      state.role.canAdmin = !!(item && item.PuedeAdministrar);
    } catch {
      state.role.canEdit = false;
      state.role.canAdmin = false;
    }
  }

  async function upsertRoleForUser(user) {
    const fieldsMap = await spHttp.getListFields(CONFIG.listNames.roles);
    const userField = resolveFieldInternal(fieldsMap, "Usuario", "User", "Persona");
    if (!userField) return;
    const filter = `${userField}Id eq ${user.Id}`;
    const url = spHttp.listItemsUrl(CONFIG.listNames.roles, `$select=Id&$filter=${encodeURIComponent(filter)}&$top=1`);
    const data = await spHttp.get(url);
    const existing = (data.value || [])[0];
    const payload = {};
    payload[`${userField}Id`] = user.Id;
    if (fieldsMap.PuedeEditar) payload.PuedeEditar = true;
    if (existing) {
      const safe = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
      const body = await withMetadata(CONFIG.listNames.roles, safe);
      await spHttp.merge(`${spHttp.listItemsUrl(CONFIG.listNames.roles)}(${existing.Id})`, body);
    } else {
      const safe = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
      await spHttp.postListItem(CONFIG.listNames.roles, safe);
    }
  }

  async function loadEditRequests() {
    const list = el("admin-req-list");
    const count = el("admin-req-count");
    if (!list || !count) return;
    try {
      const fieldsMap = await spHttp.getListFields(CONFIG.listNames.editRequests);
      const userField = resolveFieldInternal(fieldsMap, "Usuario", "User", "Persona");
      if (!userField) {
        list.innerHTML = "";
        count.textContent = "Lista sin campo Usuario.";
        return;
      }
      const selectParts = ["Id", "Title", "Estado", "FechaSolicitud", `${userField}/Title`, `${userField}/Id`];
      if (fieldsMap.EvaluacionId) selectParts.push("EvaluacionId");
      if (fieldsMap.ScenarioKey) selectParts.push("ScenarioKey");
      if (fieldsMap.TargetValor) selectParts.push("TargetValor");
      const url = spHttp.listItemsUrl(CONFIG.listNames.editRequests, `$select=${selectParts.join(",")}&$expand=${userField}&$orderby=Created desc&$top=200`);
      const data = await spHttp.get(url);
      const items = data.value || [];
      count.textContent = items.length ? `Solicitudes: ${items.length}` : "Sin solicitudes";
      list.innerHTML = items.map(i => {
        const userTitle = i[userField] && i[userField].Title ? i[userField].Title : "-";
        const estado = i.Estado || "Pendiente";
        const fecha = i.FechaSolicitud ? new Date(i.FechaSolicitud).toLocaleDateString() : "-";
        const evalTitle = i.Title || (i.EvaluacionId ? `Evaluación #${i.EvaluacionId}` : "Evaluación");
        const scenario = i.ScenarioKey || i.TargetValor || "-";
        return `<div class="item">
          <div class="inline" style="justify-content:space-between; width:100%;">
            <div>
              <strong>${userTitle}</strong>
              <div class="muted">${evalTitle} · ${scenario}</div>
              <div class="muted">${estado} · ${fecha}</div>
            </div>
            <div class="inline">
              <button class="btn" data-approve="${i.Id}" data-user="${i[userField] ? i[userField].Id : ""}">Aprobar</button>
              <button class="btn" data-reject="${i.Id}">Rechazar</button>
            </div>
          </div>
        </div>`;
      }).join("");

      list.querySelectorAll("button[data-approve]").forEach(btn => {
        btn.addEventListener("click", async () => {
          const id = Number(btn.dataset.approve);
          const userId = Number(btn.dataset.user);
          if (!id || !userId) return;
          await approveEditRequest(id, userId);
          await loadEditRequests();
          await loadUserRoles();
        });
      });
      list.querySelectorAll("button[data-reject]").forEach(btn => {
        btn.addEventListener("click", async () => {
          const id = Number(btn.dataset.reject);
          if (!id) return;
          await rejectEditRequest(id);
          await loadEditRequests();
        });
      });
    } catch (err) {
      count.textContent = "No se pudo cargar solicitudes.";
      list.innerHTML = "";
    }
  }

  async function approveEditRequest(id, userId) {
    const fieldsMap = await spHttp.getListFields(CONFIG.listNames.editRequests);
    const userPayload = { Estado: "Aprobado", FechaRespuesta: new Date().toISOString() };
    const safe = spHttp.coercePayloadByFieldTypes(userPayload, fieldsMap);
    const body = await withMetadata(CONFIG.listNames.editRequests, safe);
    await spHttp.merge(`${spHttp.listItemsUrl(CONFIG.listNames.editRequests)}(${id})`, body);
    await upsertRoleForUser({ Id: userId });
  }

  async function rejectEditRequest(id) {
    const fieldsMap = await spHttp.getListFields(CONFIG.listNames.editRequests);
    const payload = { Estado: "Rechazado", FechaRespuesta: new Date().toISOString() };
    const safe = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
    const body = await withMetadata(CONFIG.listNames.editRequests, safe);
    await spHttp.merge(`${spHttp.listItemsUrl(CONFIG.listNames.editRequests)}(${id})`, body);
  }

  async function submitEditRequest() {
    const fieldsMap = await spHttp.getListFields(CONFIG.listNames.editRequests);
    const userField = resolveFieldInternal(fieldsMap, "Usuario", "User", "Persona");
    if (!userField) return swalError("Configuración", "La lista de solicitudes no tiene campo Usuario.");
    const payload = {};
    payload[`${userField}Id`] = state.user.Id;
    payload.Estado = "Pendiente";
    payload.FechaSolicitud = new Date().toISOString();
    if (state.currentEvalId && fieldsMap.EvaluacionId) payload.EvaluacionId = state.currentEvalId;
    if (state.currentEvalItem && fieldsMap.ScenarioKey) payload.ScenarioKey = state.currentEvalItem.ScenarioKey || "";
    if (state.currentEvalItem && fieldsMap.TargetValor) payload.TargetValor = state.currentEvalItem.TargetValor || "";
    if (state.currentEvalItem && state.currentEvalItem.Title) payload.Title = state.currentEvalItem.Title;
    const safe = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
    await spHttp.postListItem(CONFIG.listNames.editRequests, safe);
    swalSuccess("Solicitud enviada", "Se envió la solicitud de edición.");
  }

  function canEditEvaluation(item) {
    const ownerId = (item.Evaluador && item.Evaluador.Id) || (item.Author && item.Author.Id) || item.EvaluadorId || item.AuthorId;
    return state.role.canEdit && ownerId && state.user && ownerId === state.user.Id;
  }

  function updateEvalEditButtons(item) {
    const reqBtn = el("eval-request-edit");
    const editBtn = el("eval-edit-start");
    if (!reqBtn || !editBtn) return;
    const canEdit = canEditEvaluation(item);
    reqBtn.classList.toggle("hidden", canEdit);
    editBtn.classList.toggle("hidden", !canEdit);
  }

  function populateEvalEditForm(item) {
    const editCard = el("eval-edit-card");
    if (!editCard) return;
    CONFIG.preguntas.forEach(key => {
      const sel = editCard.querySelector(`select[data-key="${key}"]`);
      if (sel) sel.value = getQuestionValue(item, key) ?? "";
    });
    const evento = el("eval-edit-evento");
    if (evento) evento.value = item.EventoAnalizadoSnapshot || "";
    const just = el("eval-edit-justificacion");
    if (just) just.value = item.Justificacion || "";
  }

  function resolveHierarchy() {
    const { nivel, fleet, proceso, egi, equipo } = state.wizard;
    if (nivel === "Equipo" && equipo) {
      return {
        EquipNo: equipo.Title,
        Fleet: equipo.Fleet || "",
        ProcesoSistema: equipo.ProcesoSistema || "",
        EGI: equipo.EGI || ""
      };
    }
    if (nivel === "EGI") {
      return { EquipNo: "", Fleet: fleet || "", ProcesoSistema: proceso || "", EGI: egi || "" };
    }
    if (nivel === "ProcesoSistema") {
      return { EquipNo: "", Fleet: fleet || "", ProcesoSistema: proceso || "", EGI: "" };
    }
    if (nivel === "Fleet") {
      return { EquipNo: "", Fleet: fleet || "", ProcesoSistema: "", EGI: "" };
    }
    return { EquipNo: "", Fleet: "", ProcesoSistema: "", EGI: "" };
  }

  function applyTargetFilter(items) {
    const h = resolveHierarchy();
    const fleet = state.wizard.fleet || h.Fleet;
    const proceso = state.wizard.proceso || h.ProcesoSistema;
    const egi = state.wizard.egi || h.EGI;
    return items.filter(i => {
      if (state.wizard.nivel === "Equipo") return (i.EquipNo || i.Title) === h.EquipNo;
      if (fleet && i.Fleet !== fleet) return false;
      if (proceso && i.ProcesoSistema !== proceso) return false;
      if (egi && i.EGI !== egi) return false;
      return true;
    });
  }

  function updateCascadeOptions() {
    const fleet = el("sel-fleet").value;
    const proceso = el("sel-proceso").value;
    const egi = el("sel-egi").value;
    let filtered = state.equipos;
    if (fleet) filtered = filtered.filter(i => i.Fleet === fleet);
    if (proceso) filtered = filtered.filter(i => i.ProcesoSistema === proceso);
    if (egi) filtered = filtered.filter(i => i.EGI === egi);
    const distinctProceso = [...new Set(filtered.map(i => i.ProcesoSistema).filter(Boolean))].sort();
    const distinctEgi = [...new Set(filtered.map(i => i.EGI).filter(Boolean))].sort();
    fillSelect(el("sel-proceso"), distinctProceso, "Seleccione...");
    fillSelect(el("sel-egi"), distinctEgi, "Seleccione...");
    el("sel-fleet").value = fleet;
    if (distinctProceso.includes(proceso)) el("sel-proceso").value = proceso;
    if (distinctEgi.includes(egi)) el("sel-egi").value = egi;
  }

  function setWizardState() {
    state.wizard.nivel = el("nivel-select").value;
    state.wizard.fleet = el("sel-fleet").value;
    state.wizard.proceso = el("sel-proceso").value;
    state.wizard.egi = el("sel-egi").value;
    updateTargetSummary();
  }

  function updateTargetSummary() {
    const box = el("target-summary");
    if (!box) return;
    const target = getTarget();
    if (!target.tipo) {
      box.textContent = "";
      return;
    }
    if (!target.valor) {
      box.textContent = `Estás evaluando por: ${target.tipo}`;
      return;
    }
    box.textContent = `Estás evaluando por: ${target.tipo} - ${target.valor}`;
  }

  function buildEvalEquiposItems(equipos) {
    return (equipos || []).map(eq => {
      const equipNo = typeof eq === "string" ? eq : (eq.EquipNo || eq.Title || "");
      const norm = normalizeEquip(equipNo);
      const e = state.equiposMap.get(norm) || state.equiposMap.get(equipNo) || {};
      return {
        EquipNo: e.Title || equipNo || "-",
        NombreEquipo: e.NombreEquipo || "-",
        Fleet: e.Fleet || "-",
        ProcesoSistema: e.ProcesoSistema || "-",
        EGI: e.EGI || "-"
      };
    });
  }

  function applyEvalEquiposFilters(resetPage = true) {
    const term = (el("eval-filter-equipo")?.value || "").trim().toLowerCase();
    const fleet = el("eval-filter-fleet")?.value || "";
    const proceso = el("eval-filter-proceso")?.value || "";
    const egi = el("eval-filter-egi")?.value || "";
    let filtered = state.evalDetalle.items;
    if (term) {
      filtered = filtered.filter(r =>
        String(r.EquipNo).toLowerCase().includes(term) ||
        String(r.NombreEquipo).toLowerCase().includes(term)
      );
    }
    if (fleet) filtered = filtered.filter(r => r.Fleet === fleet);
    if (proceso) filtered = filtered.filter(r => r.ProcesoSistema === proceso);
    if (egi) filtered = filtered.filter(r => r.EGI === egi);
    state.evalDetalle.filtered = filtered;
    if (resetPage) state.evalDetalle.page = 0;
    renderEvalEquiposTable();
  }

  function buildMisEvalCoverage(item) {
    let equipos = [];
    const json = item.EquiposCubiertosJson || item.EquiposCubiertosJSON || null;
    if (json) {
      try { equipos = JSON.parse(json) || []; } catch { equipos = []; }
    }
    if (!equipos.length) {
      if (item.EquipNo) equipos = [item.EquipNo];
      else if (item.TargetTipo && item.TargetValor) {
        const isDi = (e) => String(e.Estado || "").trim().toUpperCase() === "DI";
        if (item.TargetTipo === "FLEET") equipos = state.equipos.filter(e => isDi(e) && normalizeDim(e.Fleet) === normalizeDim(item.TargetValor)).map(e => e.Title);
        if (item.TargetTipo === "PROCESO_SISTEMA") equipos = state.equipos.filter(e => isDi(e) && normalizeDim(e.ProcesoSistema) === normalizeDim(item.TargetValor)).map(e => e.Title);
        if (item.TargetTipo === "EGI") equipos = state.equipos.filter(e => isDi(e) && normalizeDim(e.EGI) === normalizeDim(item.TargetValor)).map(e => e.Title);
      }
    }
    const detalles = (equipos || []).map(eq => {
      const norm = normalizeEquip(eq);
      const e = state.equiposMap.get(norm) || state.equiposMap.get(eq) || {};
      return {
        EquipNo: e.Title || eq || "-",
        NombreEquipo: e.NombreEquipo || "-",
        Fleet: e.Fleet || "-",
        ProcesoSistema: e.ProcesoSistema || "-",
        EGI: e.EGI || "-",
        BranchGerencia: e.BranchGerencia || "-",
        Estado: e.Estado || ""
      };
    });
    return {
      equipos: detalles,
      fleets: new Set(detalles.map(d => d.Fleet).filter(v => v && v !== "-")),
      procesos: new Set(detalles.map(d => d.ProcesoSistema).filter(v => v && v !== "-")),
      egis: new Set(detalles.map(d => d.EGI).filter(v => v && v !== "-"))
    };
  }

  function getMisCoverageSets(items) {
    const fleets = new Set();
    const procesos = new Set();
    const egis = new Set();
    (items || []).forEach(item => {
      if (!item._coverage) item._coverage = buildMisEvalCoverage(item);
      item._coverage.fleets.forEach(v => fleets.add(v));
      item._coverage.procesos.forEach(v => procesos.add(v));
      item._coverage.egis.forEach(v => egis.add(v));
    });
    return {
      fleets: Array.from(fleets).sort(),
      procesos: Array.from(procesos).sort(),
      egis: Array.from(egis).sort()
    };
  }

  function updateMisFilterOptions(sourceItems) {
    const prevFleet = el("mis-fleet")?.value || "";
    const prevProceso = el("mis-proceso")?.value || "";
    const prevEgi = el("mis-egi")?.value || "";
    const sets = getMisCoverageSets(sourceItems);
    fillSelect(el("mis-fleet"), sets.fleets, "Todos");
    fillSelect(el("mis-proceso"), sets.procesos, "Todos");
    fillSelect(el("mis-egi"), sets.egis, "Todos");
    if (sets.fleets.includes(prevFleet)) el("mis-fleet").value = prevFleet;
    if (sets.procesos.includes(prevProceso)) el("mis-proceso").value = prevProceso;
    if (sets.egis.includes(prevEgi)) el("mis-egi").value = prevEgi;
    return {
      fleetChanged: prevFleet && el("mis-fleet").value !== prevFleet,
      procesoChanged: prevProceso && el("mis-proceso").value !== prevProceso,
      egiChanged: prevEgi && el("mis-egi").value !== prevEgi
    };
  }

  function applyMisEvaluacionesFilters(resetPage = true) {
    const desde = el("mis-desde")?.value || "";
    const hasta = el("mis-hasta")?.value || "";
    const fleet = el("mis-fleet")?.value || "";
    const proceso = el("mis-proceso")?.value || "";
    const egi = el("mis-egi")?.value || "";
    const equipoTerm = (el("mis-equipo")?.value || "").trim().toLowerCase();

    let filtered = state.misEval.items.slice();

    if (desde) {
      const d = new Date(desde);
      filtered = filtered.filter(i => i.FechaEvaluacion && new Date(i.FechaEvaluacion) >= d);
    }
    if (hasta) {
      const h = new Date(hasta); h.setDate(h.getDate() + 1);
      filtered = filtered.filter(i => i.FechaEvaluacion && new Date(i.FechaEvaluacion) < h);
    }

    filtered = filtered.filter(item => {
      if (!item._coverage) item._coverage = buildMisEvalCoverage(item);
      const cov = item._coverage;
      if (fleet && !cov.fleets.has(fleet)) return false;
      if (proceso && !cov.procesos.has(proceso)) return false;
      if (egi && !cov.egis.has(egi)) return false;
      if (equipoTerm) {
        const match = cov.equipos.some(e =>
          String(e.EquipNo).toLowerCase().includes(equipoTerm) ||
          String(e.NombreEquipo).toLowerCase().includes(equipoTerm)
        );
        if (!match) return false;
      }
      return true;
    });

    state.misEval.filtered = filtered;
    if (resetPage) state.misEval.page = 0;

    if (!state.misEval.updatingOptions) {
      state.misEval.updatingOptions = true;
      const source = filtered.length ? filtered : state.misEval.items;
      const changed = updateMisFilterOptions(source);
      state.misEval.updatingOptions = false;
      if (changed.fleetChanged || changed.procesoChanged || changed.egiChanged) {
        applyMisEvaluacionesFilters(true);
        return;
      }
    }

    renderMisEvaluaciones();
  }

  function renderEvalEquiposTable() {
    const countEl = el("eval-equipos-count");
    const table = el("eval-equipos-table");
    const pageInput = el("eval-equipos-page");
    const pagesEl = el("eval-equipos-pages");
    const prevBtn = el("eval-equipos-prev");
    const nextBtn = el("eval-equipos-next");
    if (!table || !countEl || !pageInput || !pagesEl || !prevBtn || !nextBtn) return;

    const total = state.evalDetalle.filtered.length;
    const pageSize = state.evalDetalle.pageSize;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const page = Math.min(Math.max(0, state.evalDetalle.page), pages - 1);
    state.evalDetalle.page = page;
    const start = page * pageSize;
    const slice = state.evalDetalle.filtered.slice(start, start + pageSize);

    countEl.textContent = total ? `Mostrando ${slice.length} de ${total}` : "Sin datos de equipos.";
    pageInput.value = total ? String(page + 1) : "1";
    pagesEl.textContent = `/ ${pages}`;
    prevBtn.disabled = page <= 0;
    nextBtn.disabled = page >= pages - 1 || !total;

    if (!total) {
      table.innerHTML = "";
      return;
    }

    const rows = slice.map(r => (
      `<tr><td>${r.EquipNo}</td><td>${r.NombreEquipo}</td><td>${fmt(r.Fleet)}</td><td>${fmt(r.ProcesoSistema)}</td><td>${fmt(r.EGI)}</td></tr>`
    )).join("");
    table.innerHTML = `<div class="table-wrap"><table class="table simple"><thead><tr><th>EquipNo</th><th>Nombre</th><th>Fleet</th><th>Proceso</th><th>EGI</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  }

  function applyNivelFilters() {
    const nivel = state.wizard.nivel;
    const fleetSel = el("sel-fleet");
    const procesoSel = el("sel-proceso");
    const egiSel = el("sel-egi");

    if (nivel === "Fleet") {
      fleetSel.disabled = false;
      procesoSel.disabled = true;
      egiSel.disabled = true;
      procesoSel.value = "";
      egiSel.value = "";
    } else if (nivel === "ProcesoSistema") {
      fleetSel.disabled = true;
      procesoSel.disabled = false;
      egiSel.disabled = true;
      fleetSel.value = "";
      egiSel.value = "";
    } else if (nivel === "EGI") {
      fleetSel.disabled = true;
      procesoSel.disabled = true;
      egiSel.disabled = false;
      fleetSel.value = "";
      procesoSel.value = "";
    } else {
      fleetSel.disabled = true;
      procesoSel.disabled = true;
      egiSel.disabled = true;
      fleetSel.value = "";
      procesoSel.value = "";
      egiSel.value = "";
    }
  }

  function enableStep3() {
    const nivel = state.wizard.nivel;
    const ok = (nivel === "Fleet" && state.wizard.fleet)
      || (nivel === "ProcesoSistema" && state.wizard.proceso)
      || (nivel === "EGI" && state.wizard.egi)
      || (nivel === "Equipo" && state.wizard.equipo);
    el("to-step-3").disabled = !ok || (nivel === "Equipo" && state.wizard.existingEvalId);
  }

  function clearExistingEvaluacionNotice() {
    const box = el("equipo-existente");
    if (box) box.classList.add("hidden");
    const info = el("equipo-existente-info");
    if (info) info.textContent = "-";
    state.wizard.existingEvalId = null;
  }

  function showExistingEvaluacionNotice(item) {
    const box = el("equipo-existente");
    if (!box) return;
    const info = el("equipo-existente-info");
    const fecha = item.FechaEvaluacion ? new Date(item.FechaEvaluacion).toLocaleDateString() : "-";
    const targetTipo = item.TargetTipo || "";
    const targetValor = item.TargetValor || "";
    const targetTxt = targetTipo && targetValor ? `${targetTipo} - ${targetValor}` : "";
    let headline = "Este equipo ya fue evaluado";
    if (state.wizard.nivel === "Fleet") headline = "Esta flota ya fue evaluada";
    else if (state.wizard.nivel === "ProcesoSistema") headline = "Este proceso ya fue evaluado";
    else if (state.wizard.nivel === "EGI") headline = "Este EGI ya fue evaluado";
    box.querySelector("strong").textContent = headline;
    info.textContent = `#${item.Id} · ${item.Title || "Evaluación"} · ${fecha}${targetTxt ? ` · ${targetTxt}` : ""}`;
    box.classList.remove("hidden");
    state.wizard.existingEvalId = item.Id;
    enableStep3();
  }

  async function loadEvaluacionesJsonIndex(force = false) {
    const now = Date.now();
    if (!force && state.evalJsonIndex.loadedAt && now - state.evalJsonIndex.loadedAt < 5 * 60 * 1000) {
      return state.evalJsonIndex;
    }
    const fieldsMap = await getEvaluacionesFields();
    if (!fieldsMap.EquiposCubiertosJson) return state.evalJsonIndex;
    const selectParts = ["Id", "Title", "FechaEvaluacion", "EquiposCubiertosJson", "TargetTipo", "TargetValor", "ScenarioKey"];
    if (fieldsMap.FleetsCubiertosJson) selectParts.push("FleetsCubiertosJson");
    if (fieldsMap.ProcesosCubiertosJson) selectParts.push("ProcesosCubiertosJson");
    if (fieldsMap.EGIsCubiertosJson) selectParts.push("EGIsCubiertosJson");
    if (fieldsMap.Clasificacion) selectParts.push("Clasificacion");
    const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}&$orderby=FechaEvaluacion desc&$top=2000`);
    const items = await spHttp.fetchPaged(url);

    const byEquip = new Map();
    const equipSet = new Set();
    items.forEach(item => {
      if (!item.EquiposCubiertosJson) return;
      let equipos = [];
      try { equipos = JSON.parse(item.EquiposCubiertosJson) || []; } catch { equipos = []; }
      equipos.forEach(eq => {
        const key = normalizeEquip(eq);
        if (!key) return;
        equipSet.add(key);
        if (!byEquip.has(key)) byEquip.set(key, []);
        byEquip.get(key).push(item);
      });
    });

    state.evalJsonIndex = { loadedAt: now, items, byEquip, equipSet };
    return state.evalJsonIndex;
  }

  async function findEvaluacionByEquipJson(equipNo, force = false) {
    if (!equipNo) return null;
    const idx = await loadEvaluacionesJsonIndex(force);
    const list = idx.byEquip.get(normalizeEquip(equipNo)) || [];
    if (!list.length) return null;
    return list.reduce((best, cur) => {
      const bd = best && best.FechaEvaluacion ? new Date(best.FechaEvaluacion).getTime() : 0;
      const cd = cur && cur.FechaEvaluacion ? new Date(cur.FechaEvaluacion).getTime() : 0;
      return cd >= bd ? cur : best;
    }, list[0]);
  }

  function parseJsonList(text) {
    if (!text) return [];
    try { return JSON.parse(text) || []; } catch { return []; }
  }

  async function findEvaluacionByHierarchyJson(fleet, proceso, egi, force = false) {
    const idx = await loadEvaluacionesJsonIndex(force);
    if (!idx.items.length) return null;
    const fleetKey = normalizeEquip(fleet);
    const procesoKey = normalizeEquip(proceso);
    const egiKey = normalizeEquip(egi);
    const matches = idx.items.filter(item => {
      const fleets = parseJsonList(item.FleetsCubiertosJson).map(normalizeEquip);
      const procesos = parseJsonList(item.ProcesosCubiertosJson).map(normalizeEquip);
      const egis = parseJsonList(item.EGIsCubiertosJson).map(normalizeEquip);
      if (fleetKey && fleets.length && !fleets.includes(fleetKey)) return false;
      if (procesoKey && procesos.length && !procesos.includes(procesoKey)) return false;
      if (egiKey && egis.length && !egis.includes(egiKey)) return false;
      return (fleetKey || procesoKey || egiKey) ? true : false;
    });
    if (!matches.length) return null;
    return matches.reduce((best, cur) => {
      const bd = best && best.FechaEvaluacion ? new Date(best.FechaEvaluacion).getTime() : 0;
      const cd = cur && cur.FechaEvaluacion ? new Date(cur.FechaEvaluacion).getTime() : 0;
      return cd >= bd ? cur : best;
    }, matches[0]);
  }

  function getEquiposByFilters(fleet, proceso, egi) {
    let list = state.equipos;
    if (fleet) list = list.filter(e => e.Fleet === fleet);
    if (proceso) list = list.filter(e => e.ProcesoSistema === proceso);
    if (egi) list = list.filter(e => e.EGI === egi);
    return list.map(e => normalizeEquip(e.Title));
  }

  async function isProcesoFullyEvaluated(proceso, fleet, force = false) {
    if (!proceso) return false;
    const idx = await loadEvaluacionesJsonIndex(force);
    if (!idx.equipSet.size) return false;
    const equipos = getEquiposByFilters(fleet, proceso, null);
    if (!equipos.length) return false;
    return equipos.every(eq => idx.equipSet.has(eq));
  }

  async function isEgiFullyEvaluated(egi, fleet, proceso, force = false) {
    if (!egi) return false;
    const idx = await loadEvaluacionesJsonIndex(force);
    if (!idx.equipSet.size) return false;
    const equipos = getEquiposByFilters(fleet, proceso, egi);
    if (!equipos.length) return false;
    return equipos.every(eq => idx.equipSet.has(eq));
  }

  async function isFleetFullyEvaluated(fleet, force = false) {
    if (!fleet) return false;
    const idx = await loadEvaluacionesJsonIndex(force);
    if (!idx.equipSet.size) return false;
    const equipos = getEquiposByFilters(fleet, null, null);
    if (!equipos.length) return false;
    return equipos.every(eq => idx.equipSet.has(eq));
  }

  function findLatestEvaluacionForSet(equipSet, items) {
    if (!equipSet || !equipSet.size) return null;
    const matches = (items || []).filter(item => {
      if (!item.EquiposCubiertosJson) return false;
      try {
        const arr = JSON.parse(item.EquiposCubiertosJson) || [];
        return arr.some(eq => equipSet.has(eq));
      } catch {
        return false;
      }
    });
    if (!matches.length) return null;
    return matches.reduce((best, cur) => {
      const bd = best && best.FechaEvaluacion ? new Date(best.FechaEvaluacion).getTime() : 0;
      const cd = cur && cur.FechaEvaluacion ? new Date(cur.FechaEvaluacion).getTime() : 0;
      return cd >= bd ? cur : best;
    }, matches[0]);
  }

  async function findLatestEvaluacionForProceso(proceso, fleet, force = false) {
    if (!proceso) return null;
    const idx = await loadEvaluacionesJsonIndex(force);
    if (!idx.items.length) return null;
    const equipos = getEquiposByFilters(fleet, proceso, null);
    return findLatestEvaluacionForSet(new Set(equipos), idx.items);
  }

  async function findLatestEvaluacionForEgi(egi, fleet, proceso, force = false) {
    if (!egi) return null;
    const idx = await loadEvaluacionesJsonIndex(force);
    if (!idx.items.length) return null;
    const equipos = getEquiposByFilters(fleet, proceso, egi);
    return findLatestEvaluacionForSet(new Set(equipos), idx.items);
  }

  async function findLatestEvaluacionForFleet(fleet, force = false) {
    if (!fleet) return null;
    const idx = await loadEvaluacionesJsonIndex(force);
    if (!idx.items.length) return null;
    const equipos = getEquiposByFilters(fleet, null, null);
    return findLatestEvaluacionForSet(new Set(equipos), idx.items);
  }

  function mapTargetTipoForList(tipo) {
    if (tipo === "ProcesoSistema") return "PROCESO_SISTEMA";
    if (tipo === "Equipo") return "EQUIPO";
    return (tipo || "").toUpperCase();
  }

  async function checkEvaluacionExistenteTarget(tipo, valor) {
    if (!tipo || !valor) return null;
    const fieldsMap = await getEvaluacionesFields();
    if (!fieldsMap.TargetTipo || !fieldsMap.TargetValor) return null;
    const targetTipo = mapTargetTipoForList(tipo);
    const filter = `$filter=${encodeURIComponent(`TargetTipo eq '${escapeOData(targetTipo)}' and TargetValor eq '${escapeOData(valor)}'`)}`;
    const selectParts = ["Id", "Title", "FechaEvaluacion", "TargetTipo", "TargetValor"];
    if (fieldsMap.ScenarioKey) selectParts.push("ScenarioKey");
    const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}&$orderby=FechaEvaluacion desc&$top=1&${filter}`);
    const data = await spHttp.get(url);
    return (data.value || [])[0] || null;
  }

  async function checkEvaluacionExistenteParaEquipo(item) {
    if (!item) return null;
    const fieldsMap = await getEvaluacionesFields();
    if (!fieldsMap.TargetTipo || !fieldsMap.TargetValor) return null;
    const filters = [];
    if (item.Title) filters.push(`(TargetTipo eq 'EQUIPO' and TargetValor eq '${escapeOData(item.Title)}')`);
    if (item.EGI) filters.push(`(TargetTipo eq 'EGI' and TargetValor eq '${escapeOData(item.EGI)}')`);
    if (item.ProcesoSistema) filters.push(`(TargetTipo eq 'PROCESO_SISTEMA' and TargetValor eq '${escapeOData(item.ProcesoSistema)}')`);
    if (item.Fleet) filters.push(`(TargetTipo eq 'FLEET' and TargetValor eq '${escapeOData(item.Fleet)}')`);
    if (!filters.length) return null;
    const filter = `$filter=${encodeURIComponent(filters.join(" or "))}`;
    const selectParts = ["Id", "Title", "FechaEvaluacion", "TargetTipo", "TargetValor"];
    if (fieldsMap.ScenarioKey) selectParts.push("ScenarioKey");
    const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}&$orderby=FechaEvaluacion desc&$top=1&${filter}`);
    const data = await spHttp.get(url);
    return (data.value || [])[0] || null;
  }

  async function checkEvaluacionExistenteEquipNo(equipNo) {
    if (!equipNo) return null;
    const fieldsMap = await getEvaluacionesFields();
    if (!fieldsMap.EquipNo) return null;
    const filter = `$filter=${encodeURIComponent(`EquipNo eq '${escapeOData(equipNo)}'`)}`;
    const selectParts = ["Id", "Title", "FechaEvaluacion", "EquipNo"];
    if (fieldsMap.TargetTipo) selectParts.push("TargetTipo");
    if (fieldsMap.TargetValor) selectParts.push("TargetValor");
    if (fieldsMap.ScenarioKey) selectParts.push("ScenarioKey");
    const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}&$orderby=FechaEvaluacion desc&$top=1&${filter}`);
    const data = await spHttp.get(url);
    return (data.value || [])[0] || null;
  }

  function renderEquipoResults() {
    const container = el("equipo-results");
    container.innerHTML = "";
    const start = state.search.page * state.search.pageSize;
    const end = start + state.search.pageSize;
    const pageItems = state.search.results.slice(start, end);
    el("equipo-count").textContent = `Mostrando ${pageItems.length} de ${state.search.results.length}`;
    const totalPages = Math.max(1, Math.ceil(state.search.results.length / state.search.pageSize));
    el("equipo-page").value = String(state.search.page + 1);
    el("equipo-page-total").textContent = `/ ${totalPages}`;
    el("equipo-prev").disabled = state.search.page === 0;
    el("equipo-next").disabled = end >= state.search.results.length;
    const isSelectable = state.wizard.nivel === "Equipo";
    pageItems.forEach(item => {
      const card = document.createElement("div");
      const selected = state.wizard.equipo && state.wizard.equipo.Title === item.Title;
      const iconText = (item.NombreEquipo || item.Title || "").toString().substring(0, 1).toUpperCase();
      card.className = "item equip-item" + (isSelectable ? " selectable" : "") + (selected && isSelectable ? " selected" : "");
      card.innerHTML = `
        <div class="equip-row">
          <div class="equip-icon">${iconText}</div>
          <div>
            <div class="equip-title">${item.Title} - ${item.NombreEquipo || ""}</div>
            <div class="muted">${fmt(item.Fleet)} | ${fmt(item.ProcesoSistema)} | ${fmt(item.EGI)}</div>
          </div>
        </div>`;
      if (isSelectable) {
        card.addEventListener("click", async () => {
          state.wizard.equipo = item;
          clearExistingEvaluacionNotice();
          enableStep3();
          renderEquipoResults();
          const existente = await findEvaluacionByEquipJson(item.Title, true) || await checkEvaluacionExistenteParaEquipo(item);
          if (existente) {
            showExistingEvaluacionNotice(existente);
            return;
          }
          if (state.wizard.nivel === "ProcesoSistema") {
            const proceso = item.ProcesoSistema || state.wizard.proceso;
            const fleet = item.Fleet || state.wizard.fleet;
            if (await isProcesoFullyEvaluated(proceso, fleet, true)) {
              const lastEval = await findLatestEvaluacionForProceso(proceso, fleet, true);
              if (lastEval) showExistingEvaluacionNotice(lastEval);
            }
          } else if (state.wizard.nivel === "EGI") {
            const egi = item.EGI || state.wizard.egi;
            const fleet = item.Fleet || state.wizard.fleet;
            const proceso = item.ProcesoSistema || state.wizard.proceso;
            if (await isEgiFullyEvaluated(egi, fleet, proceso, true)) {
              const lastEval = await findLatestEvaluacionForEgi(egi, fleet, proceso, true);
              if (lastEval) showExistingEvaluacionNotice(lastEval);
            }
          } else if (state.wizard.nivel === "Fleet") {
            const fleet = item.Fleet || state.wizard.fleet;
            if (await isFleetFullyEvaluated(fleet, true)) {
              const lastEval = await findLatestEvaluacionForFleet(fleet, true);
              if (lastEval) showExistingEvaluacionNotice(lastEval);
            }
          }
        });
      }
      container.appendChild(card);
    });
  }

  async function loadEvaluacionDetalle(id) {
    const fieldsMap = await getEvaluacionesFields();
    const fieldsAll = await spHttp.getListFieldsAll(CONFIG.listNames.evaluaciones);
    let fieldsDetailed = await getEvaluacionesFieldsDetailed();
    if (!fieldsDetailed.length) {
      fieldsDetailed = Object.keys(fieldsAll || {}).map(k => ({ InternalName: k, Title: k, Hidden: false }));
    }
    const selectSet = new Set(["Id", "Title", "FechaEvaluacion"]);
    const expandSet = new Set();
    const addField = (f) => {
      const type = fieldsMap[f];
      if (!type) return;
      if (type === "User" || type === "Lookup" || type === "UserMulti" || type === "LookupMulti") {
        selectSet.add(`${f}/Title`);
        selectSet.add(`${f}/Id`);
        expandSet.add(f);
      } else {
        selectSet.add(f);
      }
    };

    [
      "TargetTipo", "TargetValor", "ScenarioKey", "EventoAnalizadoSnapshot",
      "EquiposCubiertosCount", "EquiposCubiertosJson", "VersionMetodologia",
      "Smax", "Impacto", "NC", "Clasificacion", "Justificacion",
      "EquipNo", "Fleet", "ProcesoSistema", "EGI",
      "EvaluadoPor", "Evaluador", "Author", "EvaluadorId", "AuthorId"
    ].forEach(addField);

    const qMap = resolveQuestionFieldsFromDetails(fieldsDetailed);
    Object.values(qMap).filter(Boolean).forEach(f => {
      const type = fieldsAll[f];
      if (type === "User" || type === "Lookup" || type === "UserMulti" || type === "LookupMulti") {
        selectSet.add(`${f}/Title`);
        selectSet.add(`${f}/Id`);
        expandSet.add(f);
      } else {
        selectSet.add(f);
      }
    });

    const selectParts = Array.from(selectSet);
    const expand = expandSet.size ? `&$expand=${Array.from(expandSet).join(",")}` : "";
    const base = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones);
    const url = `${base}(${id})?$select=${selectParts.join(",")}${expand}`;
    const item = await spHttp.get(url);
    state.currentEvalId = id;
    state.currentEvalItem = item;
    item._qMap = qMap;

    if (getQuestionValue(item, "FF") === null) {
      try {
        const fvat = await getFieldValuesAsText(CONFIG.listNames.evaluaciones, id);
        const normalize = (v) => String(v || "")
          .toLowerCase()
          .normalize("NFD")
          .replace(/\p{Diacritic}/gu, "")
          .replace(/[^a-z0-9]/g, "");
        const ffField = fieldsDetailed.find(f => {
          const t = normalize(f.Title);
          const n = normalize(f.InternalName);
          return t === "ff" || n === "ff" || t.includes("frecuencia") || n.includes("frecuencia") || t.includes("falla") || n.includes("falla");
        });
        const ffInternal = (qMap.FF || (ffField && ffField.InternalName)) || null;
        if (fvat && ffInternal && fvat[ffInternal] !== undefined) {
          item._ffValue = fvat[ffInternal];
        } else if (fvat) {
          const ffKey = Object.keys(fvat).find(k => {
            const nk = normalize(k);
            return nk === "ff" || nk.includes("frecuencia") || nk.includes("falla");
          });
          if (ffKey) item._ffValue = fvat[ffKey];
        }
      } catch {
        item._ffValue = item._ffValue;
      }
    }

    el("eval-detalle-title").textContent = item.Title || `Evaluación #${item.Id}`;
    const evaluadoPor = (item.EvaluadoPor && item.EvaluadoPor.Title) ? item.EvaluadoPor.Title
      : (item.EvaluadoPor || (item.Evaluador && item.Evaluador.Title) || (item.Author && item.Author.Title) || "-");
    const fecha = item.FechaEvaluacion ? new Date(item.FechaEvaluacion).toLocaleString() : "-";

    const foValue = Number(getQuestionValue(item, "FO") ?? item.FO ?? 0);
    const smaxValue = Number(item.Smax ?? 0);
    const severityValue = foValue && smaxValue
      ? Math.min(5, Math.max(1, 5 - Math.floor((foValue * smaxValue) / 5)))
      : null;

    const meta = el("eval-detalle-meta");
    meta.innerHTML = `
      <div class="summary-card">
        <div class="summary-grid summary-grid-3">
          <div class="summary-item"><div class="summary-label">Evaluado por</div><div class="summary-value">${evaluadoPor}</div></div>
          <div class="summary-item"><div class="summary-label">Fecha</div><div class="summary-value">${fecha}</div></div>
          <div class="summary-item"><div class="summary-label">Versión metodología</div><div class="summary-value">${item.VersionMetodologia || "-"}</div></div>
        </div>
        <div class="summary-grid">
          <div class="summary-item"><div class="summary-label">Target</div><div class="summary-value">${item.TargetTipo || "-"} - ${item.TargetValor || "-"}</div></div>
          <div class="summary-item"><div class="summary-label">Escenario</div><div class="summary-value">${item.ScenarioKey || "-"}</div></div>
          <div class="summary-item"><div class="summary-label">Equipos cubiertos</div><div class="summary-value">${item.EquiposCubiertosCount ?? "-"}</div></div>
        </div>
        <div class="summary-block"><div class="summary-label">Evento</div><div class="summary-value">${item.EventoAnalizadoSnapshot || "-"}</div></div>
        <div class="summary-metrics">
          <div class="metric-card"><div class="metric-title">Nivel de severidad</div><div class="metric-value">${severityValue ?? "-"}</div></div>
          <div class="metric-card"><div class="metric-title">Impacto</div><div class="metric-value">${item.Impacto ?? "-"}</div></div>
          <div class="metric-card"><div class="metric-title">NC</div><div class="metric-value">${item.NC ?? "-"}</div></div>
          <div class="metric-card"><div class="metric-title">Criticidad</div><div class="metric-value">${item.Clasificacion || "-"}</div></div>
        </div>
        <div class="summary-block">
          <div class="summary-label">Respuestas</div>
          <div class="summary-answers">
            ${CONFIG.preguntas.map(k => `<span class="answer-pill">${k}: ${getQuestionValue(item, k) ?? "-"}</span>`).join("")}
          </div>
        </div>
        <div class="summary-block"><div class="summary-label">Justificación</div><div class="summary-value">${item.Justificacion || "-"}</div></div>
      </div>`;

    let equipos = [];
    if (item.EquiposCubiertosJson) {
      try { equipos = JSON.parse(item.EquiposCubiertosJson) || []; } catch { equipos = []; }
    }
    if (!equipos.length) {
      if (item.EquipNo) equipos = [item.EquipNo];
      else if (item.TargetTipo && item.TargetValor) {
        if (item.TargetTipo === "FLEET") equipos = state.equipos.filter(e => e.Fleet === item.TargetValor).map(e => e.Title);
        if (item.TargetTipo === "PROCESO_SISTEMA") equipos = state.equipos.filter(e => e.ProcesoSistema === item.TargetValor).map(e => e.Title);
        if (item.TargetTipo === "EGI") equipos = state.equipos.filter(e => e.EGI === item.TargetValor).map(e => e.Title);
      }
    }

    updateEvalEditButtons(item);
    populateEvalEditForm(item);
    el("eval-edit-card").classList.add("hidden");

    const countEl = el("eval-equipos-count");
    countEl.textContent = equipos.length ? `Total: ${equipos.length}` : "Sin datos de equipos.";

    state.evalDetalle.items = buildEvalEquiposItems(equipos);
    state.evalDetalle.filtered = state.evalDetalle.items.slice();
    state.evalDetalle.page = 0;

    const fleets = [...new Set(state.evalDetalle.items.map(i => i.Fleet).filter(v => v && v !== "-"))].sort();
    const procesos = [...new Set(state.evalDetalle.items.map(i => i.ProcesoSistema).filter(v => v && v !== "-"))].sort();
    const egis = [...new Set(state.evalDetalle.items.map(i => i.EGI).filter(v => v && v !== "-"))].sort();
    fillSelect(el("eval-filter-fleet"), fleets, "Todos");
    fillSelect(el("eval-filter-proceso"), procesos, "Todos");
    fillSelect(el("eval-filter-egi"), egis, "Todos");
    const equipoInput = el("eval-filter-equipo");
    if (equipoInput) equipoInput.value = "";
    applyEvalEquiposFilters(true);
  }

  function searchEquipoLocal(term) {
    const t = (term || "").toLowerCase();
    const fleet = el("sel-fleet").value;
    const proceso = el("sel-proceso").value;
    const egi = el("sel-egi").value;
    let base = state.equipos;
    if (fleet) base = base.filter(i => i.Fleet === fleet);
    if (proceso) base = base.filter(i => i.ProcesoSistema === proceso);
    if (egi) base = base.filter(i => i.EGI === egi);
    if (!t || state.wizard.nivel !== "Equipo") {
      state.search.results = base;
    } else {
      state.search.results = base.filter(i =>
        (i.Title && i.Title.toLowerCase().includes(t)) ||
        (i.NombreEquipo && i.NombreEquipo.toLowerCase().includes(t))
      );
    }
    state.search.page = 0;
    renderEquipoResults();
  }

  async function loadDashboard() {
    setLoading(true);
    try {
      ["card-flotas", "card-equipos", "card-procesos", "card-egis"].forEach(id => {
        const card = el(id);
        if (card) card.classList.remove("metric-danger");
      });
      const now = new Date();
      const hoy = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const d7 = new Date(hoy); d7.setDate(d7.getDate() - 6);
      const d30 = new Date(hoy); d30.setDate(d30.getDate() - 29);
      const inicioMes = new Date(now.getFullYear(), now.getMonth(), 1);

      const select = "Id,EquipNo,EGI,Fleet,ProcesoSistema,FechaEvaluacion";
      const filter = `FechaEvaluacion ge datetime'${d30.toISOString()}'`;
      const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${select}&$filter=${encodeURIComponent(filter)}&$top=${CONFIG.pageSize}`);
      const items = await spHttp.fetchPaged(url);
      const idx = await loadEvaluacionesJsonIndex();

    const countHoy = items.filter(i => new Date(i.FechaEvaluacion) >= hoy).length;
    const count7 = items.filter(i => new Date(i.FechaEvaluacion) >= d7).length;
    const countMes = items.filter(i => new Date(i.FechaEvaluacion) >= inicioMes).length;

    const isDi = (e) => String(e.Estado || "").trim().toUpperCase() === "DI";
    const equiposDi = state.equipos.filter(isDi);
    const totalFleets = new Set(equiposDi.map(e => normalizeDim(e.Fleet)).filter(v => v && v !== "-")).size;
    const totalProcesos = new Set(equiposDi.map(e => normalizeDim(e.ProcesoSistema)).filter(v => v && v !== "-")).size;
    const totalEgi = new Set(equiposDi.map(e => normalizeDim(e.EGI)).filter(v => v && v !== "-")).size;
    const totalEquipos = new Set(equiposDi.map(e => normalizeEquip(e.Title)).filter(Boolean)).size;

    const cov = { fleets: new Set(), procesos: new Set(), egis: new Set(), equipos: new Set() };
    (idx.items || []).forEach(item => {
      const coverage = buildMisEvalCoverage(item);
      (coverage.equipos || []).forEach(eq => {
        if (String(eq.Estado || "").trim().toUpperCase() !== "DI") return;
        const eqKey = normalizeEquip(eq.EquipNo);
        if (eqKey) cov.equipos.add(eqKey);
        const f = normalizeDim(eq.Fleet);
        const p = normalizeDim(eq.ProcesoSistema);
        const g = normalizeDim(eq.EGI);
        if (f) cov.fleets.add(f);
        if (p) cov.procesos.add(p);
        if (g) cov.egis.add(g);
      });
    });

    const coveredFleets = cov.fleets.size;
    const coveredProcesos = cov.procesos.size;
    const coveredEgi = cov.egis.size;
    const coveredEquipos = cov.equipos.size;

    const goalPct = 90;

    const pct = (covered, total) => total ? (covered / total) * 100 : 0;
    const setText = (id, value) => { const node = el(id); if (node) node.textContent = value; };
    const setStatus = (id, pctValue) => {
      const card = el(id);
      if (!card) return;
      card.classList.toggle("metric-danger", pctValue < goalPct);
    };
    const setPctRowStatus = (pctId, pctValue) => {
      const node = el(pctId);
      if (!node) return;
      const row = node.closest(".metric-row");
      if (!row) return;
      row.classList.toggle("danger", pctValue < goalPct);
    };

    setAnimated("kpi-hoy", coveredFleets || 0, v => `${Math.round(v)}`);
    setAnimated("kpi-equipos-cubiertos", coveredEquipos || 0, v => `${Math.round(v)}`);
    setAnimated("kpi-procesos-ok", coveredProcesos || 0, v => `${Math.round(v)}`);
    setAnimated("kpi-egi-ok", coveredEgi || 0, v => `${Math.round(v)}`);
    setText("kpi-flotas-total", `de ${totalFleets}`);
    setText("kpi-equipos-total", `de ${totalEquipos}`);
    setText("kpi-procesos-total", `de ${totalProcesos}`);
    setText("kpi-egi-total", `de ${totalEgi}`);
    const pctFleets = pct(coveredFleets, totalFleets);
    const pctEquipos = pct(coveredEquipos, totalEquipos);
    const pctProcesos = pct(coveredProcesos, totalProcesos);
    const pctEgi = pct(coveredEgi, totalEgi);
    state.dashboardSummary = {
      totalFleets,
      totalProcesos,
      totalEgi,
      totalEquipos,
      coveredFleets,
      coveredProcesos,
      coveredEgi,
      coveredEquipos,
      pctFleets,
      pctProcesos,
      pctEgi,
      pctEquipos,
      pendingFleets: Math.max(0, totalFleets - coveredFleets),
      pendingProcesos: Math.max(0, totalProcesos - coveredProcesos),
      pendingEgi: Math.max(0, totalEgi - coveredEgi),
      pendingEquipos: Math.max(0, totalEquipos - coveredEquipos)
    };
    setAnimated("kpi-flotas-pct", pctFleets, v => `▲ ${v.toFixed(2)}%`);
    setAnimated("kpi-equipos-pct", pctEquipos, v => `▲ ${v.toFixed(2)}%`);
    setAnimated("kpi-procesos-pct", pctProcesos, v => `▲ ${v.toFixed(2)}%`);
    setAnimated("kpi-egi-pct", pctEgi, v => `▲ ${v.toFixed(2)}%`);
    setAnimated("kpi-flotas-pend", Math.max(0, totalFleets - coveredFleets), v => `▼ ${Math.round(v)}`);
    setAnimated("kpi-equipos-pend", Math.max(0, totalEquipos - coveredEquipos), v => `▼ ${Math.round(v)}`);
    setAnimated("kpi-procesos-pend", Math.max(0, totalProcesos - coveredProcesos), v => `▼ ${Math.round(v)}`);
    setAnimated("kpi-egi-pend", Math.max(0, totalEgi - coveredEgi), v => `▼ ${Math.round(v)}`);
    setStatus("card-flotas", pctFleets);
    setStatus("card-equipos", pctEquipos);
    setStatus("card-procesos", pctProcesos);
    setStatus("card-egis", pctEgi);
    updateAdvancedKpis(state.dashboardSummary);
    setPctRowStatus("kpi-flotas-pct", pctFleets);
    setPctRowStatus("kpi-equipos-pct", pctEquipos);
    setPctRowStatus("kpi-procesos-pct", pctProcesos);
    setPctRowStatus("kpi-egi-pct", pctEgi);

    const kpi7d = el("kpi-7d");
    if (kpi7d) kpi7d.textContent = count7;
    const kpiMes = el("kpi-mes");
    if (kpiMes) kpiMes.textContent = countMes;
    const kpiEq = el("kpi-equipos");
    if (kpiEq) kpiEq.textContent = totalEquipos;
    const classifCounts = { BAJA: 0, MEDIA: 0, ALTA: 0 };
    (idx.items || []).forEach(item => {
      const raw = String(item.Clasificacion || "").trim().toUpperCase();
      let clas = "BAJA";
      if (raw.includes("ALTA")) clas = "ALTA";
      else if (raw.includes("MEDIA")) clas = "MEDIA";
      else if (raw.includes("BAJA")) clas = "BAJA";
      const equipos = parseJsonList(item.EquiposCubiertosJson);
      const count = equipos.length || 0;
      if (clas === "ALTA") classifCounts.ALTA += count;
      else if (clas === "MEDIA") classifCounts.MEDIA += count;
      else classifCounts.BAJA += count;
    });
    const criticidadLegend = el("criticidad-legend");
    const criticidadCanvas = el("criticidad-chart");
    if (criticidadCanvas && window.Chart) {
      const entries = [
        { key: "BAJA", label: "Baja", value: classifCounts.BAJA, color: "#16a34a" },
        { key: "MEDIA", label: "Media", value: classifCounts.MEDIA, color: "#f59e0b" },
        { key: "ALTA", label: "Alta", value: classifCounts.ALTA, color: "#ef4444" }
      ];
      const totalEvalEquipos = entries.reduce((sum, e) => sum + e.value, 0) || 1;
      const data = entries.map(e => (e.value / totalEvalEquipos) * 100);
      const pctLabels = data.map(v => `${Math.round(v)}%`);
      if (criticidadChart) criticidadChart.destroy();
      const ctx = criticidadCanvas.getContext("2d");
      const countPlugin = {
        id: "counts",
        afterDatasetsDraw(chart) {
          const { ctx } = chart;
          chart.getDatasetMeta(0).data.forEach((arc, i) => {
            const pos = arc.tooltipPosition();
            const label = pctLabels[i];
            ctx.save();
            ctx.fillStyle = "#1f2937";
            ctx.font = "600 12px Segoe UI";
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillText(label, pos.x, pos.y);
            ctx.restore();
          });
        }
      };
      criticidadChart = new Chart(ctx, {
        type: "polarArea",
        data: {
          labels: entries.map(e => e.label),
          datasets: [{
            data,
            backgroundColor: entries.map(e => e.color + "33"),
            borderColor: entries.map(e => e.color),
            borderWidth: 1
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: { r: { ticks: { display: false }, grid: { color: "#eef2f7" } } },
          plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => `${entries[ctx.dataIndex].label}: ${entries[ctx.dataIndex].value}` } } }
        },
        plugins: [countPlugin]
      });
      if (criticidadLegend) {
        criticidadLegend.innerHTML = entries.map(e => `
          <div class="legend-item"><span class="legend-chip" style="background:${e.color}33;border-color:${e.color};"></span>Criticidad ${e.label}: ${e.value}</div>
        `).join("");
      }
    }

    const gerenciaTotals = new Map();
    const gerenciaNames = new Map();
    equiposDi.forEach(e => {
      const key = normalizeDim(e.BranchGerencia);
      if (!key) return;
      gerenciaTotals.set(key, (gerenciaTotals.get(key) || 0) + 1);
      if (!gerenciaNames.has(key)) gerenciaNames.set(key, e.BranchGerencia);
    });
    const gerenciaCovered = new Map();
    equiposDi.forEach(e => {
      const key = normalizeDim(e.BranchGerencia);
      if (!key) return;
      const equipKey = normalizeEquip(e.Title);
      if (cov.equipos.has(equipKey)) {
        gerenciaCovered.set(key, (gerenciaCovered.get(key) || 0) + 1);
      }
    });
    const gerenciaCanvas = el("gerencia-chart");
      if (gerenciaCanvas && window.Chart) {
      const rows = Array.from(gerenciaTotals.entries())
        .map(([key, total]) => {
          const covered = gerenciaCovered.get(key) || 0;
          const pct = total ? Math.min(100, Math.round((covered / total) * 100)) : 0;
          return { key, name: gerenciaNames.get(key) || key, total, covered, pct };
        })
        .sort((a, b) => b.total - a.total);
      const labels = rows.map(r => r.name);
      const values = rows.map(r => r.pct);
      const colors = rows.map(r => (r.pct >= goalPct ? "#bbf7d0" : "#fecaca"));
      if (gerenciaChart) gerenciaChart.destroy();
      const ctx = gerenciaCanvas.getContext("2d");
      const targetLine = {
        id: "target70",
        afterDraw(chart) {
          const { ctx, scales } = chart;
          const x = scales.x.getPixelForValue(goalPct);
          const yTop = scales.y.top;
          const yBottom = scales.y.bottom;
          ctx.save();
          ctx.strokeStyle = "#f59e0b";
          ctx.setLineDash([4, 6]);
          ctx.beginPath();
          ctx.moveTo(x, yTop);
          ctx.lineTo(x, yBottom);
          ctx.stroke();
          ctx.restore();
        }
      };
      const valueLabels = {
        id: "valueLabels",
        afterDatasetsDraw(chart) {
          const { ctx } = chart;
          const meta = chart.getDatasetMeta(0);
          if (!meta || !meta.data) return;
          ctx.save();
          ctx.fillStyle = "#111827";
          ctx.font = "600 11px Segoe UI";
          ctx.textAlign = "left";
          ctx.textBaseline = "middle";
          meta.data.forEach((bar, i) => {
            const v = values[i] ?? 0;
            const pos = bar.tooltipPosition();
            ctx.fillText(`${v}%`, pos.x + 6, pos.y);
          });
          ctx.restore();
        }
      };
        gerenciaChart = new Chart(ctx, {
        type: "bar",
        data: {
          labels,
          datasets: [{
            label: "% cubierto",
            data: values,
            backgroundColor: colors,
            borderRadius: 8,
            barThickness: 12,
            borderSkipped: false,
            order: 0
          }, {
            label: "Pendiente",
            data: values.map(v => Math.max(0, 100 - v)),
            backgroundColor: "#eef2f7",
            borderRadius: 8,
            barThickness: 12,
            borderSkipped: false,
            order: 1
          }]
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: { min: 0, max: 100, stacked: true, grid: { color: "#eef2f7" }, ticks: { callback: v => `${v}%` } },
            y: { stacked: true, grid: { display: false }, ticks: { color: "#475569" } }
          },
          plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => `${rows[ctx.dataIndex].covered} / ${rows[ctx.dataIndex].total} · ${rows[ctx.dataIndex].pct}%` } } }
        },
        plugins: [targetLine, valueLabels]
      });
    }

    const byFleet = {};
    items.forEach(i => { if (i.Fleet) byFleet[i.Fleet] = (byFleet[i.Fleet] || 0) + 1; });
    const maxFleet = Math.max(1, ...Object.values(byFleet));
    const chart = el("chart-fleet");
    if (chart) {
      chart.innerHTML = "";
      Object.keys(byFleet).sort().forEach(fleet => {
        const row = document.createElement("div");
        row.innerHTML = `<div class="inline" style="justify-content:space-between; gap:0.8rem;">
          <span>${fleet}</span><span class="badge">${byFleet[fleet]}</span></div>
          <div class="bar"><span style="width:${(byFleet[fleet] / maxFleet) * 100}%"></span></div>`;
        chart.appendChild(row);
      });
    }

    const byEgi = {};
    items.forEach(i => { if (i.EGI) byEgi[i.EGI] = (byEgi[i.EGI] || 0) + 1; });
    const topEgi = Object.entries(byEgi).sort((a, b) => b[1] - a[1]).slice(0, 5);
    const topBox = el("top-egi");
    if (topBox) {
      topBox.innerHTML = "";
      topEgi.forEach(([egi, count]) => {
        const item = document.createElement("div");
        item.className = "inline";
        item.innerHTML = `<span>${egi}</span><span class="badge">${count}</span>`;
        topBox.appendChild(item);
      });
    }

    const fieldsMap = await getEvaluacionesFields();
    const selectParts = ["Id", "Title", "FechaEvaluacion"];
    ["TargetTipo", "TargetValor", "EquiposCubiertosCount", "EquipNo", "Fleet", "EGI", "Clasificacion", "EvaluadoPor", "Evaluador"].forEach(f => { if (fieldsMap[f]) selectParts.push(f); });
    const equiposJsonField = resolveFieldInternal(fieldsMap, "EquiposCubiertosJson", "EquiposCubiertosJSON");
    if (equiposJsonField) selectParts.push(equiposJsonField);
    let expand = "";
    if (fieldsMap.EvaluadoPor) {
      if (fieldsMap.EvaluadoPor === "User" || fieldsMap.EvaluadoPor === "UserMulti") {
        selectParts.push("EvaluadoPor/Title");
        expand = "&$expand=EvaluadoPor";
      } else {
        selectParts.push("EvaluadoPor");
      }
    }
    if (fieldsMap.Evaluador) {
      selectParts.push("Evaluador/Title");
      expand = expand ? `${expand},Evaluador` : "&$expand=Evaluador";
    }
    const lastUrl = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}${expand}&$orderby=FechaEvaluacion desc&$top=50`);
    const last = await spHttp.get(lastUrl);
    state.dashEval.items = last.value || [];
    state.dashEval.page = 0;
    renderDashEvaluaciones();
    } finally {
      setLoading(false);
    }
  }

  function renderDashEvaluaciones() {
    const list = el("tabla-ultimas");
    const items = state.dashEval.items || [];
    const filters = state.dashboardAdvanced || {};
    const fGer = normalizeDim(filters.gerencia);
    const fSup = normalizeDim(filters.superintendencia);
    const fUni = normalizeDim(filters.unidad);
    const fFleet = normalizeDim(filters.fleet);
    const fProc = normalizeDim(filters.proceso);
    const fEgi = normalizeDim(filters.egi);
    const filteredItems = items.filter(item => {
      const cov = parseJsonList(item.EquiposCubiertosJson);
      const first = cov.find(Boolean);
      const key = normalizeEquip(first?.EquipNo || first?.Title || first || "");
      const equip = key ? state.equiposMap.get(key) : null;
      const ger = normalizeDim(equip?.BranchGerencia || "");
      const sup = normalizeDim(equip?.SiteSuperintendencia || "");
      const uni = normalizeDim(equip?.UnidadProceso || "");
      const fleet = normalizeDim(equip?.Fleet || item.Fleet || "");
      const proc = normalizeDim(equip?.ProcesoSistema || "");
      const egi = normalizeDim(equip?.EGI || "");
      if (fGer && !ger.includes(fGer)) return false;
      if (fSup && !sup.includes(fSup)) return false;
      if (fUni && !uni.includes(fUni)) return false;
      if (fFleet && !fleet.includes(fFleet)) return false;
      if (fProc && !proc.includes(fProc)) return false;
      if (fEgi && !egi.includes(fEgi)) return false;
      return true;
    });
    const dashCount = el("dash-eval-count");
    if (dashCount) dashCount.textContent = filteredItems.length ? `Análisis de criticidad: ${filteredItems.length}` : "Sin resultados";
    const start = state.dashEval.page * state.dashEval.pageSize;
    const end = start + state.dashEval.pageSize;
    const pageItems = filteredItems.slice(start, end);
    const totalPages = Math.max(1, Math.ceil(filteredItems.length / state.dashEval.pageSize));
    const dashPage = el("dash-eval-page");
    if (dashPage) dashPage.value = String(state.dashEval.page + 1);
    const dashTotal = el("dash-eval-page-total");
    if (dashTotal) dashTotal.textContent = `/ ${totalPages}`;
    const dashPrev = el("dash-eval-prev");
    if (dashPrev) dashPrev.disabled = state.dashEval.page === 0;
    const dashNext = el("dash-eval-next");
    if (dashNext) dashNext.disabled = end >= filteredItems.length;

    const rows = pageItems.map(item => {
      const meta = resolveEquipMetaFromItem(item);
      const tipo = item.TargetTipo || "-";
      const valor = item.TargetValor || item.EquipNo || "-";
      const cubiertos = item.EquiposCubiertosCount ?? "-";
      const criticidad = item.Clasificacion || "-";
      const lider = (item.EvaluadoPor && item.EvaluadoPor.Title) || (item.Evaluador && item.Evaluador.Title) || item.EvaluadoPor || item.Evaluador || "-";
      const fecha = item.FechaEvaluacion ? new Date(item.FechaEvaluacion).toLocaleDateString() : "-";
      return `<tr>
        <td>${meta.gerencia}</td>
        <td>${meta.superintendencia}</td>
        <td>${meta.unidadProceso}</td>
        <td>${tipo}</td>
        <td>${valor}</td>
        <td>${cubiertos}</td>
        <td>${criticidad}</td>
        <td>${lider}</td>
        <td>${fecha}</td>
        <td><button class="btn" data-view="${item.Id}">Ver</button></td>
      </tr>`;
    }).join("");

    if (!list) return;
    list.innerHTML = filteredItems.length ? `
      <div class="table-wrap">
        <table class="table simple">
          <thead>
            <tr>
              <th>Gerencia</th>
              <th>Superintendencia</th>
              <th>UnidadProceso</th>
              <th>Tipo</th>
              <th>Target</th>
              <th>Equipos</th>
              <th>Criticidad</th>
              <th>Líder operativo</th>
              <th>Fecha</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>` : `<div class="muted">No hay análisis de criticidad.</div>`;

    list.querySelectorAll("button[data-view]").forEach(btn => {
      btn.addEventListener("click", async () => {
        state.lastView = "view-dashboard-advanced";
        await loadEvaluacionDetalle(Number(btn.dataset.view));
        setView("view-evaluacion");
      });
    });
  }

  async function loadMisEvaluaciones() {
    const filters = [];
    if (state.user) filters.push(`EvaluadorId eq ${state.user.Id}`);
    const desde = el("mis-desde").value;
    const hasta = el("mis-hasta").value;
    if (desde) filters.push(`FechaEvaluacion ge datetime'${new Date(desde).toISOString()}'`);
    if (hasta) {
      const endDate = new Date(hasta); endDate.setDate(endDate.getDate() + 1);
      filters.push(`FechaEvaluacion lt datetime'${endDate.toISOString()}'`);
    }
    const filter = filters.length ? `$filter=${encodeURIComponent(filters.join(" and "))}` : "";
    const fieldsMap = await getEvaluacionesFields();
    const selectParts = ["Id", "Title", "FechaEvaluacion"];
    ["EquipNo", "EGI", "Fleet", "ProcesoSistema", "Clasificacion"].forEach(f => { if (fieldsMap[f]) selectParts.push(f); });
    ["TargetTipo", "TargetValor", "EquiposCubiertosCount"].forEach(f => { if (fieldsMap[f]) selectParts.push(f); });
    const equiposJsonField = resolveFieldInternal(fieldsMap, "EquiposCubiertosJson", "EquiposCubiertosJSON");
    if (equiposJsonField) selectParts.push(equiposJsonField);
    const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}&$orderby=FechaEvaluacion desc&$top=200&${filter}`);
    const data = await spHttp.get(url);
    state.misEval.items = data.value || [];
    state.misEval.page = 0;
    updateMisFilterOptions(state.misEval.items);
    applyMisEvaluacionesFilters(true);
  }

  async function loadAllEvaluaciones() {
    const filters = [];
    const equipoTerm = (el("all-equipo") && el("all-equipo").value.trim()) || "";
    const fleet = el("all-fleet") && el("all-fleet").value;
    const proceso = el("all-proceso") && el("all-proceso").value;
    const egi = el("all-egi") && el("all-egi").value;
    const desde = el("all-desde") && el("all-desde").value;
    const hasta = el("all-hasta") && el("all-hasta").value;
    if (desde) filters.push(`FechaEvaluacion ge datetime'${new Date(desde).toISOString()}'`);
    if (hasta) {
      const endDate = new Date(hasta); endDate.setDate(endDate.getDate() + 1);
      filters.push(`FechaEvaluacion lt datetime'${endDate.toISOString()}'`);
    }
    if (equipoTerm) {
      filters.push(`(substringof('${escapeOData(equipoTerm)}', EquipNo) or substringof('${escapeOData(equipoTerm)}', TargetValor))`);
    }
    if (fleet) filters.push(`Fleet eq '${escapeOData(fleet)}'`);
    if (proceso) filters.push(`ProcesoSistema eq '${escapeOData(proceso)}'`);
    if (egi) filters.push(`EGI eq '${escapeOData(egi)}'`);
    const filter = filters.length ? `$filter=${encodeURIComponent(filters.join(" and "))}` : "";
    const fieldsMap = await getEvaluacionesFields();
    const selectParts = ["Id", "Title", "FechaEvaluacion"];
    ["EquipNo", "EGI", "Fleet", "ProcesoSistema", "TargetTipo", "TargetValor", "EquiposCubiertosCount", "Clasificacion", "EvaluadoPor", "Evaluador"].forEach(f => { if (fieldsMap[f]) selectParts.push(f); });
    let expand = "";
    if (fieldsMap.EvaluadoPor) {
      if (fieldsMap.EvaluadoPor === "User" || fieldsMap.EvaluadoPor === "UserMulti") {
        selectParts.push("EvaluadoPor/Title");
        expand = "&$expand=EvaluadoPor";
      } else {
        selectParts.push("EvaluadoPor");
      }
    }
    if (fieldsMap.Evaluador) {
      selectParts.push("Evaluador/Title");
      expand = expand ? `${expand},Evaluador` : "&$expand=Evaluador";
    }
    const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}${expand}&$orderby=FechaEvaluacion desc&$top=200&${filter}`);
    const data = await spHttp.get(url);
    state.allEval.items = data.value || [];
    state.allEval.page = 0;
    renderAllEvaluaciones();
  }

  function syncAdvancedFilters() {
    const getVal = (id, current) => {
      const node = el(id);
      if (!node) return current || "";
      return String(node.value || "").trim();
    };
    state.dashboardAdvanced.gerencia = getVal("adv-gerencia", state.dashboardAdvanced.gerencia);
    state.dashboardAdvanced.superintendencia = getVal("adv-superintendencia", state.dashboardAdvanced.superintendencia);
    state.dashboardAdvanced.unidad = getVal("adv-unidad", state.dashboardAdvanced.unidad);
    state.dashboardAdvanced.fleet = getVal("adv-fleet", state.dashboardAdvanced.fleet);
    state.dashboardAdvanced.proceso = getVal("adv-proceso", state.dashboardAdvanced.proceso);
    state.dashboardAdvanced.egi = getVal("adv-egi", state.dashboardAdvanced.egi);
  }

  function clearAdvancedFilters() {
    ["adv-gerencia", "adv-superintendencia", "adv-unidad", "adv-fleet", "adv-proceso", "adv-egi"].forEach(id => {
      const node = el(id);
      if (node) node.value = "";
    });
    state.dashboardAdvanced.gerencia = "";
    state.dashboardAdvanced.superintendencia = "";
    state.dashboardAdvanced.unidad = "";
    state.dashboardAdvanced.fleet = "";
    state.dashboardAdvanced.proceso = "";
    state.dashboardAdvanced.egi = "";
  }

  function updateAdvancedFilterOptions(filteredEquipos) {
    const distinct = (arr) => [...new Set(arr.filter(Boolean))].sort();
    const setOptions = (id, values, current) => {
      const select = el(id);
      if (!select) return;
      const opts = [`<option value="">Todos</option>`].concat(values.map(v => `<option value="${v}">${v}</option>`));
      select.innerHTML = opts.join("");
      if (current && values.includes(current)) select.value = current;
      else select.value = "";
    };
    setOptions("adv-gerencia", distinct(filteredEquipos.map(e => normalizeDim(getDimValue(e, "GERENCIA")))), state.dashboardAdvanced.gerencia);
    setOptions("adv-superintendencia", distinct(filteredEquipos.map(e => normalizeDim(getDimValue(e, "SUPERINTENDENCIA")))), state.dashboardAdvanced.superintendencia);
    setOptions("adv-unidad", distinct(filteredEquipos.map(e => normalizeDim(getDimValue(e, "UNIDADPROCESO")))), state.dashboardAdvanced.unidad);
    setOptions("adv-fleet", distinct(filteredEquipos.map(e => normalizeDim(e.Fleet))), state.dashboardAdvanced.fleet);
    setOptions("adv-proceso", distinct(filteredEquipos.map(e => normalizeDim(e.ProcesoSistema))), state.dashboardAdvanced.proceso);
    setOptions("adv-egi", distinct(filteredEquipos.map(e => normalizeDim(e.EGI))), state.dashboardAdvanced.egi);
  }

  async function loadAdvancedDashboard() {
    const advCoverageCanvas = el("adv-coverage-chart");
    if (!advCoverageCanvas || !window.Chart) return;
    const idx = await loadEvaluacionesJsonIndex();
    const isDi = (e) => String(e.Estado || "").trim().toUpperCase() === "DI";
    const equiposDi = state.equipos.filter(isDi);
    const getEquipKey = (e) => normalizeEquip(e?.EquipNo || e?.Title || e?.NombreEquipo || "");
    syncAdvancedFilters();
    const filterGer = normalizeDim(state.dashboardAdvanced.gerencia);
    const filterSup = normalizeDim(state.dashboardAdvanced.superintendencia);
    const filterUni = normalizeDim(state.dashboardAdvanced.unidad);
    const filterFleet = normalizeDim(state.dashboardAdvanced.fleet);
    const filterProc = normalizeDim(state.dashboardAdvanced.proceso);
    const filterEgi = normalizeDim(state.dashboardAdvanced.egi);
    const hasFilters = Boolean(filterGer || filterSup || filterUni || filterFleet || filterProc || filterEgi);

    const filteredEquipos = equiposDi.filter(e => {
      const g = normalizeDim(getDimValue(e, "GERENCIA"));
      const s = normalizeDim(getDimValue(e, "SUPERINTENDENCIA"));
      const u = normalizeDim(getDimValue(e, "UNIDADPROCESO"));
      const f = normalizeDim(e.Fleet);
      const p = normalizeDim(e.ProcesoSistema);
      const eg = normalizeDim(e.EGI);
      if (filterGer && !g.includes(filterGer)) return false;
      if (filterSup && !s.includes(filterSup)) return false;
      if (filterUni && !u.includes(filterUni)) return false;
      if (filterFleet && !f.includes(filterFleet)) return false;
      if (filterProc && !p.includes(filterProc)) return false;
      if (filterEgi && !eg.includes(filterEgi)) return false;
      return true;
    });
    updateAdvancedFilterOptions(filteredEquipos);
    const filteredEquipSet = new Set(filteredEquipos.map(getEquipKey).filter(Boolean));

    const cov = new Set();
    (idx.items || []).forEach(item => {
      const equipos = parseJsonList(item.EquiposCubiertosJson);
      equipos.forEach(eq => {
        const key = normalizeEquip(eq.EquipNo || eq.Title || eq);
        if (!key) return;
        cov.add(key);
      });
    });

    const totalsByDim = new Map();
    const coveredByDim = new Map();
    filteredEquipos.forEach(e => {
      const key = normalizeDim(getDimValue(e, state.dashboardAdvanced.dim));
      if (!key) return;
      totalsByDim.set(key, (totalsByDim.get(key) || 0) + 1);
      if (cov.has(getEquipKey(e))) {
        coveredByDim.set(key, (coveredByDim.get(key) || 0) + 1);
      }
    });

    let rows = Array.from(totalsByDim.entries()).map(([key, total]) => {
      const covered = coveredByDim.get(key) || 0;
      const pct = total ? Math.round((covered / total) * 100) : 0;
      return { key, total, covered, pct };
    });
    rows = rows.sort((a, b) => b.pct - a.pct);

    const labels = rows.map(r => r.key);
    const coveredData = rows.map(r => r.pct);
    const pendingData = rows.map(r => Math.max(0, 100 - r.pct));

    if (advCoverageChart) advCoverageChart.destroy();
    const coverageHeight = Math.max(240, labels.length * 28);
    const coverageWidth = advCoverageCanvas.parentElement ? advCoverageCanvas.parentElement.clientWidth : advCoverageCanvas.clientWidth;
    advCoverageCanvas.height = coverageHeight;
    advCoverageCanvas.width = Math.max(320, coverageWidth || 320);
    advCoverageCanvas.style.height = `${coverageHeight}px`;
    advCoverageCanvas.style.width = "100%";
    const coverageLabels = {
      id: "coverageLabels",
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        if (!meta || !meta.data) return;
        const rightEdge = chart.chartArea?.right ?? 0;
        ctx.save();
        ctx.fillStyle = "#0f172a";
        ctx.font = "600 11px Segoe UI";
        ctx.textAlign = "left";
        ctx.textBaseline = "middle";
        meta.data.forEach((bar, i) => {
          const v = coveredData[i] ?? 0;
          const pos = bar.tooltipPosition();
          const label = `${v.toFixed(1)}%`;
          const idealX = pos.x + 6;
          const maxX = rightEdge - 6;
          if (idealX + 28 > maxX) {
            ctx.textAlign = "right";
            ctx.fillText(label, maxX, pos.y);
            ctx.textAlign = "left";
          } else {
            ctx.fillText(label, idealX, pos.y);
          }
        });
        ctx.restore();
      }
    };
    advCoverageChart = new Chart(advCoverageCanvas.getContext("2d"), {
      type: "bar",
      data: {
        labels,
        datasets: [
          { label: "% cubierto", data: coveredData, backgroundColor: "#bbf7d0", borderRadius: 8, barThickness: 12, borderSkipped: false },
          { label: "% pendiente", data: pendingData, backgroundColor: "#fecaca", borderRadius: 8, barThickness: 12, borderSkipped: false }
        ]
      },
      options: {
        indexAxis: "y",
        responsive: false,
        maintainAspectRatio: false,
        layout: { padding: { right: 32 } },
        scales: {
          x: { min: 0, max: 100, stacked: true, grid: { color: "#eef2f7" }, ticks: { callback: v => `${v}%` } },
          y: { stacked: true, grid: { display: false }, ticks: { color: "#475569" } }
        },
        plugins: { legend: { display: false } }
      },
      plugins: [coverageLabels]
    });

    const totalFleets = new Set(filteredEquipos.map(e => normalizeDim(e.Fleet)).filter(Boolean)).size;
    const totalProcesos = new Set(filteredEquipos.map(e => normalizeDim(e.ProcesoSistema)).filter(Boolean)).size;
    const totalEgi = new Set(filteredEquipos.map(e => normalizeDim(e.EGI)).filter(Boolean)).size;
    const totalEquipos = filteredEquipSet.size;

    const coveredFleets = new Set();
    const coveredProcesos = new Set();
    const coveredEgi = new Set();
    const coveredEquipos = new Set();
    filteredEquipos.forEach(e => {
      const key = getEquipKey(e);
      if (!key || !cov.has(key)) return;
      coveredEquipos.add(key);
      const f = normalizeDim(e.Fleet);
      const p = normalizeDim(e.ProcesoSistema);
      const g = normalizeDim(e.EGI);
      if (f) coveredFleets.add(f);
      if (p) coveredProcesos.add(p);
      if (g) coveredEgi.add(g);
    });

    const pct = (covered, total) => total ? (covered / total) * 100 : 0;
    const summary = {
      totalFleets,
      totalProcesos,
      totalEgi,
      totalEquipos,
      coveredFleets: coveredFleets.size,
      coveredProcesos: coveredProcesos.size,
      coveredEgi: coveredEgi.size,
      coveredEquipos: coveredEquipos.size,
      pctFleets: pct(coveredFleets.size, totalFleets),
      pctProcesos: pct(coveredProcesos.size, totalProcesos),
      pctEgi: pct(coveredEgi.size, totalEgi),
      pctEquipos: pct(coveredEquipos.size, totalEquipos),
      pendingFleets: Math.max(0, totalFleets - coveredFleets.size),
      pendingProcesos: Math.max(0, totalProcesos - coveredProcesos.size),
      pendingEgi: Math.max(0, totalEgi - coveredEgi.size),
      pendingEquipos: Math.max(0, totalEquipos - coveredEquipos.size)
    };
    updateAdvancedKpis(summary);

    const gerTotals = new Map();
    const gerCovered = new Map();
    const supTotals = new Map();
    const supCovered = new Map();
    const uniTotals = new Map();
    const uniCovered = new Map();

    filteredEquipos.forEach(e => {
      const key = normalizeEquip(e.Title);
      const g = normalizeDim(getDimValue(e, "GERENCIA"));
      const s = normalizeDim(getDimValue(e, "SUPERINTENDENCIA"));
      const u = normalizeDim(getDimValue(e, "UNIDADPROCESO"));
      if (g) gerTotals.set(g, (gerTotals.get(g) || 0) + 1);
      if (s) supTotals.set(s, (supTotals.get(s) || 0) + 1);
      if (u) uniTotals.set(u, (uniTotals.get(u) || 0) + 1);
      if (key && cov.has(key)) {
        if (g) gerCovered.set(g, (gerCovered.get(g) || 0) + 1);
        if (s) supCovered.set(s, (supCovered.get(s) || 0) + 1);
        if (u) uniCovered.set(u, (uniCovered.get(u) || 0) + 1);
      }
    });

    const countFullyCovered = (totals, covered) => {
      let count = 0;
      totals.forEach((total, key) => {
        if ((covered.get(key) || 0) >= total && total > 0) count += 1;
      });
      return count;
    };
    const setCovCard = (valueId, totalId, pctId, pendId, cardId, covered, total) => {
      const pctVal = total ? (covered / total) * 100 : 0;
      const valNode = el(valueId);
      const totalNode = el(totalId);
      const pctNode = el(pctId);
      const pendNode = el(pendId);
      if (valNode) valNode.textContent = `${covered}`;
      if (totalNode) totalNode.textContent = `de ${total}`;
      if (pctNode) pctNode.textContent = `▲ ${pctVal.toFixed(2)}%`;
      if (pendNode) pendNode.textContent = `▼ ${Math.max(0, total - covered)}`;
      const card = el(cardId);
      if (card) card.classList.toggle("metric-danger", pctVal < goalPct);
      if (pctNode) {
        const row = pctNode.closest(".metric-row");
        if (row) row.classList.toggle("danger", pctVal < goalPct);
      }
    };
    setCovCard(
      "adv-gerencia-cov",
      "adv-gerencia-total",
      "adv-gerencia-pct",
      "adv-gerencia-cov-pend",
      "adv-card-gerencia",
      countFullyCovered(gerTotals, gerCovered),
      gerTotals.size
    );
    setCovCard(
      "adv-super-cov",
      "adv-super-total",
      "adv-super-pct",
      "adv-super-cov-pend",
      "adv-card-super",
      countFullyCovered(supTotals, supCovered),
      supTotals.size
    );
    setCovCard(
      "adv-unidad-cov",
      "adv-unidad-total",
      "adv-unidad-pct",
      "adv-unidad-cov-pend",
      "adv-card-unidad",
      countFullyCovered(uniTotals, uniCovered),
      uniTotals.size
    );

    const classifCounts = { BAJA: 0, MEDIA: 0, ALTA: 0 };
    const altaByTarget = new Map();
    const trend = new Map();
    (idx.items || []).forEach(item => {
      const equipos = parseJsonList(item.EquiposCubiertosJson);
      let matchCount = 0;
      equipos.forEach(eq => {
        const key = normalizeEquip(eq.EquipNo || eq.Title || eq);
        if (!key) return;
        if ((hasFilters || filteredEquipSet.size) && !filteredEquipSet.has(key)) return;
        matchCount += 1;
      });
      if (!matchCount) return;
      const raw = String(item.Clasificacion || "").trim().toUpperCase();
      let clas = "BAJA";
      if (raw.includes("ALTA")) clas = "ALTA";
      else if (raw.includes("MEDIA")) clas = "MEDIA";
      classifCounts[clas] += matchCount;
      if (clas === "ALTA") {
        const target = item.TargetValor || item.EquipNo || "-";
        altaByTarget.set(target, (altaByTarget.get(target) || 0) + matchCount);
      }
      const fecha = item.FechaEvaluacion ? new Date(item.FechaEvaluacion) : null;
      if (fecha && !Number.isNaN(fecha.getTime())) {
        const key = `${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2, "0")}`;
        trend.set(key, (trend.get(key) || 0) + 1);
      }
    });

    const advCriticidadCanvas = el("adv-criticidad-chart");
    if (advCriticidadCanvas) {
      const entries = [
        { label: "Baja", value: classifCounts.BAJA, color: "#16a34a" },
        { label: "Media", value: classifCounts.MEDIA, color: "#f59e0b" },
        { label: "Alta", value: classifCounts.ALTA, color: "#ef4444" }
      ];
      const total = entries.reduce((sum, e) => sum + e.value, 0) || 1;
      const data = entries.map(e => (e.value / total) * 100);
      if (advCriticidadChart) advCriticidadChart.destroy();
      const criticidadCounts = {
        id: "criticidadCounts",
        afterDatasetsDraw(chart) {
          const { ctx } = chart;
          const meta = chart.getDatasetMeta(0);
          if (!meta || !meta.data) return;
          ctx.save();
          ctx.fillStyle = "#1f2937";
          ctx.font = "600 12px Segoe UI";
          ctx.textAlign = "center";
          ctx.textBaseline = "middle";
          meta.data.forEach((arc, i) => {
            const pos = arc.tooltipPosition();
            const label = entries[i]?.value ?? 0;
            ctx.fillText(String(label), pos.x, pos.y);
          });
          ctx.restore();
        }
      };
      advCriticidadChart = new Chart(advCriticidadCanvas.getContext("2d"), {
        type: "polarArea",
        data: {
          labels: entries.map(e => e.label),
          datasets: [{ data, backgroundColor: entries.map(e => e.color + "33"), borderColor: entries.map(e => e.color), borderWidth: 1 }]
        },
        options: { responsive: true, maintainAspectRatio: false, scales: { r: { ticks: { display: false }, grid: { color: "#eef2f7" } } }, plugins: { legend: { display: false } } },
        plugins: [criticidadCounts]
      });
      const legend = el("adv-criticidad-legend");
      if (legend) {
        legend.innerHTML = entries.map(e => `
          <div class="legend-item"><span class="legend-chip" style="background:${e.color}33;border-color:${e.color};"></span>Criticidad ${e.label}: ${e.value}</div>
        `).join("");
      }
    }

    const advAltaCanvas = el("adv-alta-chart");
    if (advAltaCanvas) {
      const topAlta = Array.from(altaByTarget.entries()).sort((a, b) => b[1] - a[1]);
      const altaLabels = topAlta.map(i => i[0]);
      const altaValues = topAlta.map(i => i[1]);
      if (advAltaChart) advAltaChart.destroy();
      const altaHeight = Math.max(240, altaLabels.length * 28);
      const altaWidth = advAltaCanvas.parentElement ? advAltaCanvas.parentElement.clientWidth : advAltaCanvas.clientWidth;
      advAltaCanvas.height = altaHeight;
      advAltaCanvas.width = Math.max(320, altaWidth || 320);
      advAltaCanvas.style.height = `${altaHeight}px`;
      advAltaCanvas.style.width = "100%";
      const altaLabelsPlugin = {
        id: "altaLabels",
        afterDatasetsDraw(chart) {
          const { ctx } = chart;
          const meta = chart.getDatasetMeta(0);
          if (!meta || !meta.data) return;
          ctx.save();
          ctx.fillStyle = "#1f2937";
          ctx.font = "600 11px Segoe UI";
          ctx.textAlign = "left";
          ctx.textBaseline = "middle";
          meta.data.forEach((bar, i) => {
            const v = altaValues[i] ?? 0;
            const pos = bar.tooltipPosition();
            ctx.fillText(String(v), pos.x + 6, pos.y);
          });
          ctx.restore();
        }
      };
      advAltaChart = new Chart(advAltaCanvas.getContext("2d"), {
        type: "bar",
        data: { labels: altaLabels, datasets: [{ data: altaValues, backgroundColor: "#fecaca", borderRadius: 8, barThickness: 12 }] },
        options: { indexAxis: "y", responsive: false, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { grid: { color: "#eef2f7" } }, y: { grid: { display: false } } } },
        plugins: [altaLabelsPlugin]
      });
    }

    const advTrendCanvas = el("adv-trend-chart");
    if (advTrendCanvas) {
      const trendLabels = Array.from(trend.keys()).sort();
      const trendValues = trendLabels.map(k => trend.get(k) || 0);
      if (advTrendChart) advTrendChart.destroy();
      advTrendChart = new Chart(advTrendCanvas.getContext("2d"), {
        type: "line",
        data: { labels: trendLabels, datasets: [{ data: trendValues, borderColor: "#60a5fa", backgroundColor: "#93c5fd33", fill: true, tension: 0.3 }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { grid: { color: "#eef2f7" } } } }
      });
    }

    state.dashEval.page = 0;
    renderDashEvaluaciones();
  }

  function renderAllEvaluaciones() {
    const list = el("all-eval-list");
    const items = state.allEval.items || [];
    const countEl = el("all-eval-count");
    if (countEl) countEl.textContent = items.length ? `Análisis de criticidad: ${items.length}` : "Sin resultados";
    const start = state.allEval.page * state.allEval.pageSize;
    const end = start + state.allEval.pageSize;
    const pageItems = items.slice(start, end);
    const totalPages = Math.max(1, Math.ceil(items.length / state.allEval.pageSize));
    const pageInput = el("all-eval-page");
    if (pageInput) pageInput.value = String(state.allEval.page + 1);
    const pageTotal = el("all-eval-page-total");
    if (pageTotal) pageTotal.textContent = `/ ${totalPages}`;
    const prevBtn = el("all-eval-prev");
    const nextBtn = el("all-eval-next");
    if (prevBtn) prevBtn.disabled = state.allEval.page === 0;
    if (nextBtn) nextBtn.disabled = end >= items.length;

    const rows = pageItems.map(item => {
      const meta = resolveEquipMetaFromItem(item);
      const lider = (item.EvaluadoPor && item.EvaluadoPor.Title) || (item.Evaluador && item.Evaluador.Title) || item.EvaluadoPor || item.Evaluador || "-";
      const fecha = item.FechaEvaluacion ? new Date(item.FechaEvaluacion).toLocaleDateString() : "-";
      const tipo = item.TargetTipo || "-";
      const valor = item.TargetValor || item.EquipNo || "-";
      const cubiertos = item.EquiposCubiertosCount ?? "-";
      const clasif = item.Clasificacion || "-";
      return `<tr>
        <td>${meta.gerencia}</td>
        <td>${meta.superintendencia}</td>
        <td>${meta.unidadProceso}</td>
        <td>${tipo}</td>
        <td>${valor}</td>
        <td>${cubiertos}</td>
        <td>${clasif}</td>
        <td>${lider}</td>
        <td>${fecha}</td>
        <td><button class="btn" data-view="${item.Id}">Ver</button></td>
      </tr>`;
    }).join("");

    if (!list) return;
    list.innerHTML = items.length ? `
      <div class="table-wrap">
        <table class="table simple">
          <thead>
            <tr>
              <th>Gerencia</th>
              <th>Superintendencia</th>
              <th>UnidadProceso</th>
              <th>Tipo</th>
              <th>Target</th>
              <th>Equipos</th>
              <th>Criticidad</th>
              <th>LíderOperativo</th>
              <th>Fecha</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${rows || ""}</tbody>
        </table>
      </div>` : `<div class="muted">No hay análisis de criticidad.</div>`;

    list.querySelectorAll("button[data-view]").forEach(btn => {
      btn.addEventListener("click", async () => {
        state.lastView = "view-all";
        await loadEvaluacionDetalle(Number(btn.dataset.view));
        setView("view-evaluacion");
      });
    });
  }

  function renderMisEvaluaciones() {
    const list = el("mis-list");
    const items = state.misEval.filtered || state.misEval.items || [];
    el("mis-eval-count").textContent = items.length ? `Análisis de criticidad: ${items.length}` : "Sin resultados";
    const start = state.misEval.page * state.misEval.pageSize;
    const end = start + state.misEval.pageSize;
    const pageItems = items.slice(start, end);
    const totalPages = Math.max(1, Math.ceil(items.length / state.misEval.pageSize));
    el("mis-eval-page").value = String(state.misEval.page + 1);
    el("mis-eval-page-total").textContent = `/ ${totalPages}`;
    el("mis-eval-prev").disabled = state.misEval.page === 0;
    el("mis-eval-next").disabled = end >= items.length;

    const rows = pageItems.map(item => {
      const meta = resolveEquipMetaFromItem(item);
      const tipo = item.TargetTipo || "-";
      const valor = item.TargetValor || item.EquipNo || "-";
      const cubiertos = item.EquiposCubiertosCount ?? "-";
      const criticidad = item.Clasificacion || "-";
      const fecha = item.FechaEvaluacion ? new Date(item.FechaEvaluacion).toLocaleDateString() : "-";
      return `<tr>
        <td>${meta.gerencia}</td>
        <td>${meta.superintendencia}</td>
        <td>${meta.unidadProceso}</td>
        <td>${tipo}</td>
        <td>${valor}</td>
        <td>${cubiertos}</td>
        <td>${criticidad}</td>
        <td>${fecha}</td>
        <td><button class="btn" data-view="${item.Id}">Ver</button></td>
      </tr>`;
    }).join("");

    list.innerHTML = items.length ? `
      <div class="table-wrap">
        <table class="table simple">
          <thead>
            <tr>
              <th>Gerencia</th>
              <th>Superintendencia</th>
              <th>UnidadProceso</th>
              <th>Tipo</th>
              <th>Target</th>
              <th>Equipos</th>
              <th>Criticidad</th>
              <th>Fecha</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>` : `<div class="muted">No hay análisis de criticidad.</div>`;

    list.querySelectorAll("button[data-view]").forEach(btn => {
      btn.addEventListener("click", async () => {
        state.lastView = "view-mis";
        await loadEvaluacionDetalle(Number(btn.dataset.view));
        setView("view-evaluacion");
      });
    });
  }

  async function loadAdminEscenarios() {
    const fieldsMap = await spHttp.getListFields(CONFIG.listNames.escenarios);
    const escFields = {
      nivel: resolveFieldInternal(fieldsMap, "NivelAplicacion", "NivelAplicación", "Nivel"),
      activo: resolveFieldInternal(fieldsMap, "Activo", "Estado"),
      fleet: resolveFieldInternal(fieldsMap, "Fleet"),
      proceso: resolveFieldInternal(fieldsMap, "ProcesoSistema", "Proceso"),
      egi: resolveFieldInternal(fieldsMap, "EGI"),
      equipNo: resolveFieldInternal(fieldsMap, "EquipNo"),
      evento: resolveFieldInternal(fieldsMap, "EventoAnalizado", "Evento")
    };
    state.escFields = escFields;

    const filters = [];
    const nivel = el("admin-nivel").value;
    const fleet = el("admin-fleet").value;
    const proceso = el("admin-proceso").value;
    const egi = el("admin-egi").value;
    if (nivel && escFields.nivel) filters.push(`${escFields.nivel} eq '${escapeOData(nivel)}'`);
    if (fleet && escFields.fleet) filters.push(`${escFields.fleet} eq '${escapeOData(fleet)}'`);
    if (proceso && escFields.proceso) filters.push(`${escFields.proceso} eq '${escapeOData(proceso)}'`);
    if (egi && escFields.egi) filters.push(`${escFields.egi} eq '${escapeOData(egi)}'`);
    const filter = filters.length ? `$filter=${encodeURIComponent(filters.join(" and "))}` : "";

    const selectParts = ["Id", "Title"]; 
    [escFields.nivel, escFields.activo, escFields.fleet, escFields.proceso, escFields.egi, escFields.equipNo, escFields.evento]
      .filter(Boolean)
      .forEach(f => selectParts.push(f));

    const url = spHttp.listItemsUrl(CONFIG.listNames.escenarios, `$select=${selectParts.join(",")}&$orderby=Modified desc&$top=50&${filter}`);
    const data = await spHttp.get(url);
    const list = el("admin-list");
    list.innerHTML = "";
    (data.value || []).forEach(raw => {
      const item = normalizeEscenarioItem(raw);
      const row = document.createElement("div");
      row.className = "item";
      row.innerHTML = `<div class="inline" style="justify-content:space-between;">
        <div>
          <strong>${item.Title}</strong>
          <div class="muted">${item.NivelAplicacion || "-"} | ${fmt(item.Fleet)} | ${fmt(item.ProcesoSistema)} | ${fmt(item.EGI)} | ${fmt(item.EquipNo)}</div>
        </div>
        <div class="inline">
          <span class="pill">${item.Activo ? "Activo" : "Inactivo"}</span>
          <button class="btn" data-edit="${item.Id}">Editar</button>
        </div>
      </div>`;
      row.querySelector("button").addEventListener("click", () => openAdminEditor(item));
      list.appendChild(row);
    });
  }

  async function loadAdminEvaluaciones() {
    const filters = [];
    const term = el("admin-eval-term").value.trim();
    const desde = el("admin-eval-desde").value;
    const hasta = el("admin-eval-hasta").value;
    if (desde) filters.push(`FechaEvaluacion ge datetime'${new Date(desde).toISOString()}'`);
    if (hasta) {
      const endDate = new Date(hasta); endDate.setDate(endDate.getDate() + 1);
      filters.push(`FechaEvaluacion lt datetime'${endDate.toISOString()}'`);
    }
    if (term) filters.push(`(substringof('${escapeOData(term)}', Title) or substringof('${escapeOData(term)}', EquipNo))`);
    const filter = filters.length ? `$filter=${encodeURIComponent(filters.join(" and "))}` : "";
    const fieldsMap = await getEvaluacionesFields();
    const selectParts = ["Id", "Title", "FechaEvaluacion"];
    ["EquipNo", "EGI", "Fleet", "TargetTipo", "TargetValor", "EquiposCubiertosCount", "NC", "Clasificacion"].forEach(f => { if (fieldsMap[f]) selectParts.push(f); });
    if (fieldsMap.EstadoEvaluacion) selectParts.push("EstadoEvaluacion");
    let expand = "";
    if (fieldsMap.Evaluador) { selectParts.push("Evaluador/Title"); expand = "&$expand=Evaluador"; }
    if (fieldsMap.EvaluadoPor) {
      if (fieldsMap.EvaluadoPor === "User" || fieldsMap.EvaluadoPor === "UserMulti") {
        selectParts.push("EvaluadoPor/Title");
        expand = expand ? `${expand},EvaluadoPor` : "&$expand=EvaluadoPor";
      } else {
        selectParts.push("EvaluadoPor");
      }
    }
    const url = spHttp.listItemsUrl(CONFIG.listNames.evaluaciones, `$select=${selectParts.join(",")}${expand}&$orderby=FechaEvaluacion desc&$top=200&${filter}`);
    const data = await spHttp.get(url);
    state.adminEval.items = data.value || [];
    state.adminEval.page = 0;
    renderAdminEvaluaciones();
  }

  function renderAdminEvaluaciones() {
    const list = el("admin-eval-list");
    const items = state.adminEval.items || [];
    el("admin-eval-count").textContent = items.length ? `Análisis de criticidad: ${items.length}` : "Sin resultados";
    const start = state.adminEval.page * state.adminEval.pageSize;
    const end = start + state.adminEval.pageSize;
    const pageItems = items.slice(start, end);
    const totalPages = Math.max(1, Math.ceil(items.length / state.adminEval.pageSize));
    el("admin-eval-page").value = String(state.adminEval.page + 1);
    el("admin-eval-page-total").textContent = `/ ${totalPages}`;
    el("admin-eval-prev").disabled = state.adminEval.page === 0;
    el("admin-eval-next").disabled = end >= items.length;

    const rows = pageItems.map(item => {
      const meta = resolveEquipMetaFromItem(item);
      const lider = item.EvaluadoPor && item.EvaluadoPor.Title ? item.EvaluadoPor.Title : (item.EvaluadoPor || (item.Evaluador && item.Evaluador.Title) || "-");
      const fecha = item.FechaEvaluacion ? new Date(item.FechaEvaluacion).toLocaleDateString() : "-";
      const tipo = item.TargetTipo || "-";
      const valor = item.TargetValor || item.EquipNo || "-";
      const cubiertos = item.EquiposCubiertosCount ?? "-";
      const clasif = item.Clasificacion || "-";
      return `<tr>
        <td>${meta.gerencia}</td>
        <td>${meta.superintendencia}</td>
        <td>${meta.unidadProceso}</td>
        <td>${tipo}</td>
        <td>${valor}</td>
        <td>${cubiertos}</td>
        <td>${clasif}</td>
        <td>${lider}</td>
        <td>${fecha}</td>
        <td>
          <button class="btn" data-view="${item.Id}">Ver</button>
        </td>
      </tr>`;
    }).join("");

    list.innerHTML = items.length ? `
      <div class="table-wrap">
        <table class="table simple">
          <thead>
            <tr>
              <th>Gerencia</th>
              <th>Superintendencia</th>
              <th>UnidadProceso</th>
              <th>Tipo</th>
              <th>Target</th>
              <th>Equipos</th>
              <th>Criticidad</th>
              <th>LíderOperativo</th>
              <th>Fecha</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${rows || ""}</tbody>
        </table>
      </div>` : `<div class="muted">No hay analisis de criticidad con esos filtros.</div>`;

    list.querySelectorAll("button[data-view]").forEach(btn => {
      btn.addEventListener("click", async () => {
        await loadEvaluacionDetalle(Number(btn.dataset.view));
        setView("view-evaluacion");
      });
    });
  }

  function openAdminEditor(item) {
    el("admin-editor").classList.remove("hidden");
    if (item) {
      state.adminEditingId = item.Id;
      el("admin-editor-title").textContent = `Editar escenario #${item.Id}`;
      el("sc-title").value = item.Title || "";
      el("sc-activo").value = item.Activo ? "1" : "0";
      el("sc-nivel").value = item.NivelAplicacion || "";
      el("sc-fleet").value = item.Fleet || "";
      el("sc-proceso").value = item.ProcesoSistema || "";
      el("sc-egi").value = item.EGI || "";
      el("sc-equipno").value = item.EquipNo || "";
      el("sc-evento").value = item.EventoAnalizado || "";
    } else {
      state.adminEditingId = null;
      el("admin-editor-title").textContent = "Nuevo escenario";
      el("sc-title").value = "";
      el("sc-activo").value = "1";
      el("sc-nivel").value = "Corporativo";
      el("sc-fleet").value = "";
      el("sc-proceso").value = "";
      el("sc-egi").value = "";
      el("sc-equipno").value = "";
      el("sc-evento").value = "";
    }
  }

  async function saveAdminEscenario() {
    const fieldsMap = await spHttp.getListFields(CONFIG.listNames.escenarios);
    const escFields = state.escFields || {
      nivel: resolveFieldInternal(fieldsMap, "NivelAplicacion", "NivelAplicación", "Nivel"),
      activo: resolveFieldInternal(fieldsMap, "Activo", "Estado"),
      fleet: resolveFieldInternal(fieldsMap, "Fleet"),
      proceso: resolveFieldInternal(fieldsMap, "ProcesoSistema", "Proceso"),
      egi: resolveFieldInternal(fieldsMap, "EGI"),
      equipNo: resolveFieldInternal(fieldsMap, "EquipNo"),
      evento: resolveFieldInternal(fieldsMap, "EventoAnalizado", "Evento")
    };

    const payload = { Title: el("sc-title").value };
    if (escFields.activo) payload[escFields.activo] = el("sc-activo").value === "1";
    if (escFields.nivel) payload[escFields.nivel] = el("sc-nivel").value;
    if (escFields.fleet) payload[escFields.fleet] = el("sc-fleet").value;
    if (escFields.proceso) payload[escFields.proceso] = el("sc-proceso").value;
    if (escFields.egi) payload[escFields.egi] = el("sc-egi").value;
    if (escFields.equipNo) payload[escFields.equipNo] = el("sc-equipno").value;
    if (escFields.evento) payload[escFields.evento] = el("sc-evento").value;

    if (!payload.Title || (escFields.nivel && !payload[escFields.nivel]) || (escFields.evento && !payload[escFields.evento])) {
      swalInfo("Faltan datos", "Complete Title, Nivel y EventoAnalizado.");
      return;
    }
    const safe = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
    if (state.adminEditingId) {
      const url = `${spHttp.listItemsUrl(CONFIG.listNames.escenarios)}/(${state.adminEditingId})`;
      await spHttp.merge(url, safe);
    } else {
      await spHttp.postListItem(CONFIG.listNames.escenarios, safe);
    }
    await loadAdminEscenarios();
    el("admin-editor").classList.add("hidden");
  }

  function normalizeEscenarioItem(raw) {
    const f = state.escFields || {};
    return {
      Id: raw.Id,
      Title: raw.Title,
      Activo: f.activo ? !!raw[f.activo] : !!raw.Activo,
      NivelAplicacion: f.nivel ? raw[f.nivel] : raw.NivelAplicacion,
      Fleet: f.fleet ? raw[f.fleet] : raw.Fleet,
      ProcesoSistema: f.proceso ? raw[f.proceso] : raw.ProcesoSistema,
      EGI: f.egi ? raw[f.egi] : raw.EGI,
      EquipNo: f.equipNo ? raw[f.equipNo] : raw.EquipNo,
      EventoAnalizado: f.evento ? raw[f.evento] : raw.EventoAnalizado
    };
  }

  async function reabrirEvaluacion(id) {
    if (!confirm("¿Reabrir evaluación?")) return;
    const payload = {
      EstadoEvaluacion: "REABIERTA",
      ReabiertaPorId: state.user.Id,
      ReabiertaFecha: new Date().toISOString()
    };
    const fieldsMap = await getEvaluacionesFields();
    const safePayload = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
    const url = `${spHttp.listItemsUrl(CONFIG.listNames.evaluaciones)}/(${id})`;
    await spHttp.merge(url, safePayload);
    await loadAdminEvaluaciones();
  }

  async function getEvaluacionesFields() {
    if (state.evalFields) return state.evalFields;
    state.evalFields = await spHttp.getListFields(CONFIG.listNames.evaluaciones);
    return state.evalFields;
  }

  function getTarget() {
    const h = resolveHierarchy();
    const tipo = state.wizard.nivel;
    const valor = tipo === "Equipo" ? h.EquipNo : (tipo === "EGI" ? h.EGI : (tipo === "ProcesoSistema" ? h.ProcesoSistema : h.Fleet));
    return { tipo, valor };
  }

  function calcCoberturaCount() {
    const { tipo, valor } = getTarget();
    if (!valor) return 0;
    if (tipo === "Equipo") return 1;
    if (tipo === "Fleet") return state.equipos.filter(i => i.Fleet === valor).length;
    if (tipo === "ProcesoSistema") return state.equipos.filter(i => i.ProcesoSistema === valor).length;
    if (tipo === "EGI") return state.equipos.filter(i => i.EGI === valor).length;
    return 0;
  }

  function computeCoverageLists() {
    const tipo = state.wizard.nivel;
    if (tipo === "Equipo" && state.wizard.equipo && state.wizard.equipo.Title) {
      const eq = normalizeEquip(state.wizard.equipo.Title);
      return { equipos: [eq], fleets: [], procesos: [], egis: [] };
    }
    const filtered = applyTargetFilter(state.equipos);
    const equipos = filtered.map(e => normalizeEquip(e.Title));

    const fleetSet = new Set();
    const procesoSet = new Set();
    const egiSet = new Set();
    filtered.forEach(item => {
      if (item.Fleet) fleetSet.add(item.Fleet);
      if (item.ProcesoSistema) procesoSet.add(item.ProcesoSistema);
      if (item.EGI) egiSet.add(item.EGI);
    });

    let fleets = Array.from(fleetSet);
    let procesos = Array.from(procesoSet);
    let egis = Array.from(egiSet);

    if (tipo === "Equipo") {
      fleets = [];
      procesos = [];
      egis = [];
    } else if (tipo === "EGI") {
      fleets = [];
      procesos = [];
    } else if (tipo === "ProcesoSistema") {
      fleets = [];
    }

    return { equipos, fleets, procesos, egis };
  }

  function formatCoverageSummary(coverage) {
    if (!coverage) return "-";
    return `Flotas: ${coverage.fleets.length} · Procesos: ${coverage.procesos.length} · EGI: ${coverage.egis.length} · Equipos: ${coverage.equipos.length}`;
  }

  function renderCoverageSummaryStep4() {
    const box = el("coverage-summary-step4");
    if (!box) return;
    if (!state.wizard.coverage) {
      state.wizard.coverage = computeCoverageLists();
      state.wizard.equiposCubiertos = state.wizard.coverage.equipos.length;
    }
    const cov = state.wizard.coverage;
    const target = getTarget();
    if (!target.tipo) {
      box.innerHTML = "";
      return;
    }
    box.innerHTML = `
      <div class="summary-card">
        <div class="summary-grid summary-grid-3">
          <div class="summary-item"><div class="summary-label">Nivel</div><div class="summary-value">${target.tipo}</div></div>
          <div class="summary-item"><div class="summary-label">Target</div><div class="summary-value">${target.valor || "-"}</div></div>
          <div class="summary-item"><div class="summary-label">Equipos</div><div class="summary-value">${cov.equipos.length}</div></div>
          <div class="summary-item"><div class="summary-label">Flotas</div><div class="summary-value">${cov.fleets.length}</div></div>
          <div class="summary-item"><div class="summary-label">Procesos</div><div class="summary-value">${cov.procesos.length}</div></div>
          <div class="summary-item"><div class="summary-label">EGI</div><div class="summary-value">${cov.egis.length}</div></div>
        </div>
      </div>`;
  }

  async function loadScenarioCandidates(searchTerm) {
    const { tipo, valor } = getTarget();
    if (!tipo || !valor) return [];
    const term = String(searchTerm || "").trim();
    if (CONFIG.requireScenarioSearch && term.length < 3) return [];
    const items = await scenarioService.listCandidates(tipo, valor, CONFIG.maxEscenarios, term);
    const filtered = applyTargetFilter(items);
    const deduped = new Map();
    filtered.forEach(item => {
      const key = `${item.Title || ""}|${item.EventoAnalizado || ""}`;
      if (!deduped.has(key)) deduped.set(key, item);
    });
    return Array.from(deduped.values());
  }

  async function loadScenarioCandidatesBySearch(term) {
    const results = await scenarioService.searchCandidates(term, 50);
    const details = [];
    for (const r of results) {
      const item = await scenarioService.getById(r.ListItemID);
      if (item && item.Activo !== false) details.push(item);
    }
    const filtered = applyTargetFilter(details);
    const deduped = new Map();
    filtered.forEach(item => {
      const key = `${item.Title || ""}|${item.EventoAnalizado || ""}`;
      if (!deduped.has(key)) deduped.set(key, item);
    });
    return Array.from(deduped.values());
  }

  function filterScenarioCandidates(items) {
    const term = (state.wizard.escenarioSearch || "").trim().toLowerCase();
    if (!term) return items;
    return items.filter(item =>
      (item.Title && item.Title.toLowerCase().includes(term)) ||
      (item.EventoAnalizado && item.EventoAnalizado.toLowerCase().includes(term))
    );
  }

  function renderScenarioCandidates(items) {
    const list = el("escenario-list");
    list.innerHTML = "";
    const filtered = filterScenarioCandidates(items);
    const start = state.wizard.escenarioPage * state.wizard.escenarioPageSize;
    const end = start + state.wizard.escenarioPageSize;
    const pageItems = filtered.slice(start, end);
    el("escenario-count").textContent = `Escenarios: ${pageItems.length} de ${filtered.length}`;
    el("escenario-prev").disabled = state.wizard.escenarioPage === 0;
    el("escenario-next").disabled = end >= filtered.length;
    const totalPages = Math.max(1, Math.ceil(filtered.length / state.wizard.escenarioPageSize));
    el("escenario-page").value = String(state.wizard.escenarioPage + 1);
    el("escenario-page-total").textContent = `/ ${totalPages}`;

    pageItems.forEach(item => {
      const badge = (item.EquipNo || "").toString().substring(0, 12);
      const row = document.createElement("div");
      row.className = "item selectable" + (state.wizard.escenario && state.wizard.escenario.Id === item.Id ? " selected" : "");
      const preview = extractPreview(item.EventoAnalizado, 1).slice(0, 140) + (item.EventoAnalizado && item.EventoAnalizado.length > 140 ? "…" : "");
      const equipNo = (item.EquipNo || "").toString().trim();
      const titleText = (item.Title || "").toString();
      const titleDisplay = (equipNo && titleText === equipNo) ? "" : titleText;
      row.innerHTML = `<div class="inline" style="justify-content:space-between; width:100%;">
        <div>
          ${titleDisplay ? `<strong>${titleDisplay}</strong>` : ""}
          <div class="scenario-strong">${fmt(item.Fleet)} | ${fmt(item.ProcesoSistema)} | ${fmt(item.EGI)}${equipNo ? ` | ${equipNo}` : ""}</div>
          <div class="scenario-preview">${preview}</div>
        </div>
        <span class="badge badge-right">${badge}</span>
      </div>`;
      row.addEventListener("click", () => {
        state.wizard.escenario = item;
        el("escenario-meta").textContent = `#${item.Id} - ${item.Title}`;
        state.wizard.eventoAnalizado = item.EventoAnalizado || "";
        el("evento-analizado").value = state.wizard.eventoAnalizado;
        const edit = el("evento-analizado-edit");
        if (edit) edit.value = state.wizard.eventoAnalizado;
        el("to-step-4").disabled = false;
        renderScenarioCandidates(state.wizard.escenarioCandidates);
      });
      list.appendChild(row);
    });
  }

  async function guardarEvaluacion() {
    const h = resolveHierarchy();
    const escenario = state.wizard.escenario;
    const respuestas = state.wizard.respuestas;
    const justificacion = state.wizard.justificacion;
    const target = getTarget();
    const fieldsMap = await getEvaluacionesFields();
    const { smax, impacto, nc, clasificacion } = calcularResultado(respuestas);

    const coverage = computeCoverageLists();
    const equiposJson = coverage.equipos.length && coverage.equipos.length <= CONFIG.equiposJsonMax ? JSON.stringify(coverage.equipos) : null;
    const fleetsJson = coverage.fleets.length ? JSON.stringify(coverage.fleets) : null;
    const procesosJson = coverage.procesos.length ? JSON.stringify(coverage.procesos) : null;
    const egisJson = coverage.egis.length ? JSON.stringify(coverage.egis) : null;

    const payload = {
      Title: `EV-${new Date().toISOString().replace(/[-:T]/g, "").slice(0, 8)}-${target.tipo}-${target.valor}-${escenario.Title}`,
      EquipNo: h.EquipNo,
      EGI: h.EGI,
      ProcesoSistema: h.ProcesoSistema,
      Fleet: h.Fleet,
      TargetTipo: target.tipo === "ProcesoSistema" ? "PROCESO_SISTEMA" : target.tipo.toUpperCase(),
      TargetValor: target.valor,
      ScenarioKey: escenario.Title,
      EventoAnalizadoSnapshot: state.wizard.eventoAnalizado || escenario.EventoAnalizado,
      EquiposCubiertosCount: coverage.equipos.length,
      Justificacion: justificacion,
      EvaluadorId: state.user.Id,
      FechaEvaluacion: new Date().toISOString(),
      EstadoEvaluacion: "CERRADA"
    };

    const qMap = resolveQuestionFields(fieldsMap);
    CONFIG.preguntas.forEach(key => {
      const fieldName = qMap[key] || key;
      if (fieldsMap[fieldName]) payload[fieldName] = Number(respuestas[key] || 0);
    });

    if (fieldsMap.VersionMetodologia) payload.VersionMetodologia = CONFIG.versionMetodologia;
    if (fieldsMap.EvaluadoPor) {
      const evalType = fieldsMap.EvaluadoPor;
      if (evalType === "User" || evalType === "UserMulti") {
        payload.EvaluadoPorId = state.user.Id;
      } else {
        payload.EvaluadoPor = state.user.Title || state.user.Email || "";
      }
    }
    const equiposJsonField = resolveFieldInternal(fieldsMap, "EquiposCubiertosJson", "EquiposCubiertosJSON");
    const fleetsJsonField = resolveFieldInternal(fieldsMap, "FleetsCubiertosJson", "FleetCubiertosJson");
    const procesosJsonField = resolveFieldInternal(fieldsMap, "ProcesosCubiertosJson", "ProcesoCubiertosJson");
    const egisJsonField = resolveFieldInternal(fieldsMap, "EGIsCubiertosJson", "EGICubiertosJson");
    if (equiposJsonField && equiposJson) payload[equiposJsonField] = equiposJson;
    if (fleetsJsonField && fleetsJson) payload[fleetsJsonField] = fleetsJson;
    if (procesosJsonField && procesosJson) payload[procesosJsonField] = procesosJson;
    if (egisJsonField && egisJson) payload[egisJsonField] = egisJson;
    if (fieldsMap.ScenarioId || fieldsMap.EscenarioId) payload[fieldsMap.ScenarioId ? "ScenarioId" : "EscenarioId"] = escenario.Id;
    if (fieldsMap.Smax) payload.Smax = smax;
    if (fieldsMap.Impacto) payload.Impacto = impacto;
    if (fieldsMap.NC) payload.NC = nc;
    if (fieldsMap.Clasificacion) payload.Clasificacion = clasificacion;

    const safePayload = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
    await spHttp.postListItem(CONFIG.listNames.evaluaciones, safePayload);
  }

  function calcularResultado(respuestas) {
    const vals = {
      FO: Number(respuestas.FO || 0),
      FIN: Number(respuestas.FIN || 0),
      HSE: Number(respuestas.HSE || 0),
      MA: Number(respuestas.MA || 0),
      SOC: Number(respuestas.SOC || 0),
      DDHH: Number(respuestas.DDHH || 0),
      REP: Number(respuestas.REP || 0),
      LEGAL: Number(respuestas.LEGAL || 0),
      TI: Number(respuestas.TI || 0),
      FF: Number(respuestas.FF || 0)
    };
    const smax = Math.max(vals.FIN, vals.HSE, vals.MA, vals.SOC, vals.DDHH, vals.REP, vals.LEGAL, vals.TI);
    const impacto = smax * vals.FO;
    const getSeveridad = (fo, smaxLocal) => {
      const base = 5 - Math.floor((fo * smaxLocal) / 5);
      return Math.min(5, Math.max(1, base || 1));
    };
    const getCriticidadMatrix = (fo, smaxLocal, ff) => {
      const matrix = {
        1: [5, 10, 20, 35, 55],
        2: [15, 25, 40, 60, 80],
        3: [30, 45, 65, 85, 100],
        4: [50, 70, 90, 105, 115],
        5: [75, 95, 110, 120, 125]
      };
      const severity = getSeveridad(fo, smaxLocal);
      const ffIndex = Math.min(5, Math.max(1, ff)) - 1;
      const row = matrix[severity] || matrix[1];
      return row[ffIndex] ?? 0;
    };
    const severity = getSeveridad(vals.FO, smax);
    const nc = getCriticidadMatrix(vals.FO, smax, vals.FF);
    const clasificacion = nc <= 30 ? "BAJA" : (nc <= 74 ? "MEDIA" : "ALTA");
    return { smax, impacto, nc, clasificacion, severity, vals };
  }

  function buildResumen() {
    const h = resolveHierarchy();
    const escenario = state.wizard.escenario;
    const respuestas = state.wizard.respuestas;
    const resumen = el("resumen");
    const target = getTarget();
    const resultado = calcularResultado(respuestas);
    const cov = state.wizard.coverage || computeCoverageLists();
    resumen.innerHTML = `
      <div class="summary-card">
        <div class="summary-grid">
          <div class="summary-item">
            <div class="summary-label">Target</div>
            <div class="summary-value">${target.tipo} - ${target.valor}</div>
          </div>
          <div class="summary-item">
            <div class="summary-label">Equipos cubiertos</div>
            <div class="summary-value">${state.wizard.equiposCubiertos}</div>
          </div>
          <div class="summary-item">
            <div class="summary-label">Escenario</div>
            <div class="summary-value">#${escenario.Id} - ${escenario.Title}</div>
          </div>
          <div class="summary-item"><div class="summary-label">Flotas</div><div class="summary-value">${cov.fleets.length}</div></div>
          <div class="summary-item"><div class="summary-label">Procesos</div><div class="summary-value">${cov.procesos.length}</div></div>
          <div class="summary-item"><div class="summary-label">EGI</div><div class="summary-value">${cov.egis.length}</div></div>
        </div>
        <div class="summary-block">
          <div class="summary-label">Evento</div>
          <div class="summary-value">${state.wizard.eventoAnalizado || escenario.EventoAnalizado || "-"}</div>
        </div>
        <div class="summary-metrics">
          <div class="metric-card">
            <div class="metric-title">Nivel de severidad</div>
            <div class="metric-value">${resultado.severity}</div>
            <div class="metric-desc">Nivel de severidad usado en la matriz (1–5), calculado con FO y Smax.</div>
          </div>
          <div class="metric-card">
            <div class="metric-title">Impacto</div>
            <div class="metric-value">${resultado.impacto}</div>
            <div class="metric-desc">Impacto = Smax × FO</div>
          </div>
          <div class="metric-card">
            <div class="metric-title inline" style="gap:0.4rem;">
              <span>NC</span>
              <button class="icon-btn" type="button" id="open-criticidad-matrix" aria-label="Ver matriz">!</button>
            </div>
            <div class="metric-value">${resultado.nc}</div>
            <div class="metric-desc">NC se obtiene de la matriz según Severidad.</div>
          </div>
          <div class="metric-card">
            <div class="metric-title">Criticidad</div>
            <div class="metric-value">${resultado.clasificacion}</div>
            <div class="metric-desc">≤30 BAJA · ≤74 MEDIA · &gt;74 ALTA.</div>
          </div>
        </div>
        <div class="summary-block">
          <div class="summary-label">Respuestas</div>
          <div class="summary-answers">${CONFIG.preguntas.map(k => `<span class="answer-pill">${k}: ${respuestas[k] || "-"}</span>`).join("")}</div>
        </div>
        <div class="summary-block">
          <div class="summary-label">Justificación</div>
          <div class="summary-value">${state.wizard.justificacion || "-"}</div>
        </div>
      </div>`;

    const infoBtn = el("open-criticidad-matrix");
    if (infoBtn) infoBtn.addEventListener("click", () => openCriticidadMatrix({ severity: resultado.severity, ff: respuestas.FF }));
  }

  function getCriticidadImageUrl() {
    const base = (window._spPageContextInfo && window._spPageContextInfo.webServerRelativeUrl)
      ? window._spPageContextInfo.webServerRelativeUrl
      : (window.location.pathname.split("/SiteAssets/")[0] || "");
    const root = base.endsWith("/") ? base.slice(0, -1) : base;
    return `${root}/SiteAssets/WebAC/SiteAssets/AppAC/img/matrix.png`;
  }

  function openCriticidadMatrix({ severity, ff }) {
    if (window.openCriticidadMatrixDynamic) {
      window.openCriticidadMatrixDynamic({ severity, ff });
      return;
    }
    const imageUrl = getCriticidadImageUrl();
    if (window.Swal) {
      Swal.fire({
        title: "Matriz de criticidad",
        text: "",
        imageUrl,
        imageWidth: 560,
        imageHeight: 420,
        imageAlt: "Matriz de criticidad",
        confirmButtonText: "Cerrar"
      });
      return;
    }
    toggleCriticidadModal(true);
  }

  function ensureCriticidadMatrixSrc() {
    const img = el("criticidad-matrix-img");
    if (!img || img.getAttribute("src")) return;
    const base = (window._spPageContextInfo && window._spPageContextInfo.webServerRelativeUrl)
      ? window._spPageContextInfo.webServerRelativeUrl
      : (window.location.pathname.split("/SiteAssets/")[0] || "");
    const root = base.endsWith("/") ? base.slice(0, -1) : base;
    const fallbackSvg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 620 380" role="img" aria-label="Matriz de criticidad">
  <rect x="20" y="20" width="580" height="340" rx="14" fill="#ffffff" stroke="#e5e7eb" />
  <text x="310" y="48" text-anchor="middle" font-size="14" fill="#0f172a" font-weight="700">FACTOR DE FRECUENCIA DE FALLA (FF)</text>
  <text x="32" y="200" text-anchor="middle" font-size="12" fill="#0f172a" font-weight="700" transform="rotate(-90 32 200)">NIVEL DE SEVERIDAD</text>
  <g font-size="12" fill="#0f172a" font-weight="700">
    <text x="190" y="80" text-anchor="middle">1</text>
    <text x="270" y="80" text-anchor="middle">2</text>
    <text x="350" y="80" text-anchor="middle">3</text>
    <text x="430" y="80" text-anchor="middle">4</text>
    <text x="510" y="80" text-anchor="middle">5</text>
  </g>
  <g font-size="12" fill="#0f172a" font-weight="700">
    <text x="110" y="130" text-anchor="middle">5</text>
    <text x="110" y="180" text-anchor="middle">4</text>
    <text x="110" y="230" text-anchor="middle">3</text>
    <text x="110" y="280" text-anchor="middle">2</text>
    <text x="110" y="330" text-anchor="middle">1</text>
  </g>
  <g font-size="12" fill="#0f172a" font-weight="700" text-anchor="middle">
    <rect x="150" y="110" width="80" height="40" fill="#f97316" /><text x="190" y="136">75</text>
    <rect x="230" y="110" width="80" height="40" fill="#ef4444" /><text x="270" y="136">95</text>
    <rect x="310" y="110" width="80" height="40" fill="#ef4444" /><text x="350" y="136">110</text>
    <rect x="390" y="110" width="80" height="40" fill="#ef4444" /><text x="430" y="136">120</text>
    <rect x="470" y="110" width="80" height="40" fill="#ef4444" /><text x="510" y="136">125</text>
    <rect x="150" y="160" width="80" height="40" fill="#f59e0b" /><text x="190" y="186">50</text>
    <rect x="230" y="160" width="80" height="40" fill="#f59e0b" /><text x="270" y="186">70</text>
    <rect x="310" y="160" width="80" height="40" fill="#ef4444" /><text x="350" y="186">90</text>
    <rect x="390" y="160" width="80" height="40" fill="#ef4444" /><text x="430" y="186">105</text>
    <rect x="470" y="160" width="80" height="40" fill="#ef4444" /><text x="510" y="186">115</text>
    <rect x="150" y="210" width="80" height="40" fill="#22c55e" /><text x="190" y="236">30</text>
    <rect x="230" y="210" width="80" height="40" fill="#f59e0b" /><text x="270" y="236">45</text>
    <rect x="310" y="210" width="80" height="40" fill="#f59e0b" /><text x="350" y="236">65</text>
    <rect x="390" y="210" width="80" height="40" fill="#ef4444" /><text x="430" y="236">85</text>
    <rect x="470" y="210" width="80" height="40" fill="#ef4444" /><text x="510" y="236">100</text>
    <rect x="150" y="260" width="80" height="40" fill="#22c55e" /><text x="190" y="286">15</text>
    <rect x="230" y="260" width="80" height="40" fill="#22c55e" /><text x="270" y="286">25</text>
    <rect x="310" y="260" width="80" height="40" fill="#f59e0b" /><text x="350" y="286">40</text>
    <rect x="390" y="260" width="80" height="40" fill="#ef4444" /><text x="430" y="286">60</text>
    <rect x="470" y="260" width="80" height="40" fill="#ef4444" /><text x="510" y="286">80</text>
    <rect x="150" y="310" width="80" height="40" fill="#22c55e" /><text x="190" y="336">5</text>
    <rect x="230" y="310" width="80" height="40" fill="#22c55e" /><text x="270" y="336">10</text>
    <rect x="310" y="310" width="80" height="40" fill="#22c55e" /><text x="350" y="336">20</text>
    <rect x="390" y="310" width="80" height="40" fill="#f59e0b" /><text x="430" y="336">35</text>
    <rect x="470" y="310" width="80" height="40" fill="#f59e0b" /><text x="510" y="336">55</text>
  </g>
</svg>`;
    const dataUri = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(fallbackSvg)}`;
    const paths = [
      `${root}/SiteAssets/WebAC/SiteAssets/AppAC/img/matrix.png`,
      `${root}/SiteAssets/AppAC/img/matrix.png`
    ];
    let idx = 0;
    img.onerror = () => {
      idx += 1;
      if (idx < paths.length) {
        img.setAttribute("src", paths[idx]);
        return;
      }
      img.onerror = null;
      img.setAttribute("src", dataUri);
    };
    img.setAttribute("src", paths[0]);
  }

  function toggleCriticidadModal(show) {
    const modal = el("criticidad-modal");
    if (!modal) return;
    if (show) ensureCriticidadMatrixSrc();
    modal.classList.toggle("open", show);
    modal.setAttribute("aria-hidden", show ? "false" : "true");
  }
  function wireEvents() {
    const on = (id, event, handler) => {
      const node = el(id);
      if (node) node.addEventListener(event, handler);
      return node;
    };

    on("nav-dashboard", "click", async () => {
      setView("view-dashboard-advanced");
      await loadAdvancedDashboard();
    });
    on("nav-wizard", "click", () => setView("view-wizard"));
    on("nav-mis", "click", () => setView("view-mis"));
    const navAll = el("nav-all");
    if (navAll) navAll.addEventListener("click", () => setView("view-all"));
    on("nav-admin", "click", () => setView("view-admin"));

    on("refresh-dashboard", "click", async () => {
      await loadDashboard();
      await loadAdvancedDashboard();
    });
    on("refresh-mis", "click", loadMisEvaluaciones);
    const allRefresh = el("all-eval-refresh");
    if (allRefresh) allRefresh.addEventListener("click", loadAllEvaluaciones);
    ["all-desde", "all-hasta"].forEach(id => {
      const input = el(id);
      if (input) input.addEventListener("change", loadAllEvaluaciones);
    });
    const allEquipo = el("all-equipo");
    if (allEquipo) allEquipo.addEventListener("input", loadAllEvaluaciones);
    ["all-fleet", "all-proceso", "all-egi"].forEach(id => {
      const sel = el(id);
      if (sel) sel.addEventListener("change", loadAllEvaluaciones);
    });

    ["adv-gerencia", "adv-superintendencia", "adv-unidad", "adv-fleet", "adv-proceso", "adv-egi"].forEach(id => {
      const input = el(id);
      if (input) input.addEventListener("change", async () => {
        await loadAdvancedDashboard();
      });
    });
    const advClear = el("adv-clear-filters");
    if (advClear) advClear.addEventListener("click", async () => {
      clearAdvancedFilters();
      await loadAdvancedDashboard();
    });

    const tabs = document.querySelectorAll("#dash-adv-tabs .dash-tab");
    tabs.forEach(btn => {
      btn.addEventListener("click", async () => {
        tabs.forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        state.dashboardAdvanced.dim = btn.dataset.dim || "GERENCIA";
        await loadAdvancedDashboard();
      });
    });
    ["mis-fleet", "mis-proceso", "mis-egi"].forEach(id => {
      const sel = el(id);
      if (sel) sel.addEventListener("change", () => applyMisEvaluacionesFilters(true));
    });
    ["mis-desde", "mis-hasta"].forEach(id => {
      const input = el(id);
      if (input) input.addEventListener("change", () => applyMisEvaluacionesFilters(true));
    });
    const misEquipo = el("mis-equipo");
    if (misEquipo) misEquipo.addEventListener("input", () => applyMisEvaluacionesFilters(true));

    const criticidadModal = el("criticidad-modal");
    const criticidadClose = el("criticidad-modal-close");
    if (criticidadClose) criticidadClose.addEventListener("click", () => toggleCriticidadModal(false));
    if (criticidadModal) {
      criticidadModal.addEventListener("click", (e) => {
        if (e.target === criticidadModal) toggleCriticidadModal(false);
      });
    }

    if (el("view-wizard")) {
      on("nivel-select", "change", () => {
        state.wizard.nivel = el("nivel-select").value;
        const step2 = el("to-step-2");
        if (step2) step2.disabled = !state.wizard.nivel;
        resetFilters();
        applyNivelFilters();
      });

      on("to-step-2", "click", () => {
        setStep(2);
        const nivel = state.wizard.nivel;
        setWizardState();
        const equipoSearch = el("equipo-search");
        if (equipoSearch) equipoSearch.classList.remove("hidden");
        const txtEquipo = el("txt-equipo");
        const equipoPrev = el("equipo-prev");
        const equipoNext = el("equipo-next");
        const equipoPage = el("equipo-page");
        if (txtEquipo) txtEquipo.disabled = nivel !== "Equipo";
        if (equipoPrev) equipoPrev.disabled = nivel !== "Equipo";
        if (equipoNext) equipoNext.disabled = nivel !== "Equipo";
        if (equipoPage) equipoPage.disabled = nivel !== "Equipo";
        const step4 = el("to-step-4");
        if (step4) step4.disabled = true;
        applyNivelFilters();
        updateCascadeOptions();
        enableStep3();
        if (txtEquipo) searchEquipoLocal(txtEquipo.value);
      });

      on("reset-filters", "click", () => {
        resetFilters();
        updateCascadeOptions();
        searchEquipoLocal("");
      });

      on("ver-evaluacion-existente", "click", async () => {
        if (!state.wizard.existingEvalId) return;
        try {
          state.lastView = "view-wizard";
          await loadEvaluacionDetalle(state.wizard.existingEvalId);
          setView("view-evaluacion");
        } catch (err) {
          console.error(err);
          const meta = el("eval-detalle-meta");
          if (meta) {
            const msg = (err && err.message) ? err.message : "Error desconocido";
            meta.innerHTML = `<div class="alert error">No se pudo cargar la evaluación. Detalle: ${msg}</div>`;
            setView("view-evaluacion");
          } else {
            alert("No se pudo cargar la evaluación. Intente nuevamente.");
          }
        }
      });

      on("eval-back", "click", () => {
        const backView = state.lastView || "view-wizard";
        setView(backView);
        if (backView === "view-wizard") setStep(2);
      });

    const reqBtn = el("eval-request-edit");
    if (reqBtn) reqBtn.addEventListener("click", submitEditRequest);

    const editStart = el("eval-edit-start");
    if (editStart) editStart.addEventListener("click", () => {
      const card = el("eval-edit-card");
      if (card) card.classList.remove("hidden");
    });

    const editCancel = el("eval-edit-cancel");
    if (editCancel) editCancel.addEventListener("click", () => {
      const card = el("eval-edit-card");
      if (card) card.classList.add("hidden");
    });

    const editSave = el("eval-edit-save");
    if (editSave) editSave.addEventListener("click", async () => {
      const card = el("eval-edit-card");
      if (!card) return;
      const currentId = state.currentEvalId || null;
      if (!currentId) return;
      if (!canEditEvaluation(state.currentEvalItem || {})) {
        return swalError("Sin permisos", "No tiene permisos para editar esta evaluación.");
      }
      const respuestas = {};
      let ok = true;
      card.querySelectorAll("select").forEach(sel => {
        respuestas[sel.dataset.key] = sel.value;
        if (!sel.value) ok = false;
      });
      const just = el("eval-edit-justificacion").value.trim();
      if (!ok || !just) {
        return swalInfo("Faltan datos", "Complete todas las respuestas y la justificación.");
      }

      const evento = el("eval-edit-evento").value.trim();
      const { smax, impacto, nc, clasificacion } = calcularResultado(respuestas);
      const fieldsMap = await getEvaluacionesFields();
      const payload = {
        Justificacion: just,
        EventoAnalizadoSnapshot: evento,
        Smax: smax,
        Impacto: impacto,
        NC: nc,
        Clasificacion: clasificacion
      };
      const qMap = resolveQuestionFields(fieldsMap);
      CONFIG.preguntas.forEach(k => {
        const fieldName = qMap[k] || k;
        if (fieldsMap[fieldName]) payload[fieldName] = Number(respuestas[k] || 0);
      });
      const safe = spHttp.coercePayloadByFieldTypes(payload, fieldsMap);
      const body = await withMetadata(CONFIG.listNames.evaluaciones, safe);
      await spHttp.merge(`${spHttp.listItemsUrl(CONFIG.listNames.evaluaciones)}(${currentId})`, body);
      swalSuccess("Cambios guardados", "El análisis de criticidad fue actualizado con éxito.");
      await loadEvaluacionDetalle(currentId);
    });

    ["eval-filter-equipo", "eval-filter-fleet", "eval-filter-proceso", "eval-filter-egi"].forEach(id => {
      const input = el(id);
      if (!input) return;
      const evt = id === "eval-filter-equipo" ? "input" : "change";
      input.addEventListener(evt, () => applyEvalEquiposFilters(true));
    });

    const prevEval = el("eval-equipos-prev");
    const nextEval = el("eval-equipos-next");
    const pageEval = el("eval-equipos-page");
    if (prevEval && nextEval && pageEval) {
      prevEval.addEventListener("click", () => {
        state.evalDetalle.page = Math.max(0, state.evalDetalle.page - 1);
        renderEvalEquiposTable();
      });
      nextEval.addEventListener("click", () => {
        state.evalDetalle.page += 1;
        renderEvalEquiposTable();
      });
      pageEval.addEventListener("change", () => {
        const total = state.evalDetalle.filtered.length;
        const pages = Math.max(1, Math.ceil(total / state.evalDetalle.pageSize));
        const desired = parseInt(pageEval.value, 10);
        if (Number.isNaN(desired)) return;
        state.evalDetalle.page = Math.min(Math.max(0, desired - 1), pages - 1);
        renderEvalEquiposTable();
      });
    }

      on("back-step-1", "click", () => setStep(1));

      ["sel-fleet", "sel-proceso", "sel-egi"].forEach(id => {
        const node = el(id);
        if (!node) return;
        node.addEventListener("change", () => {
        setWizardState();
        updateCascadeOptions();
        enableStep3();
        searchEquipoLocal(el("txt-equipo").value);
        if (state.wizard.nivel !== "Equipo") {
          const tipo = state.wizard.nivel;
          const valor = getTarget().valor;
          if (tipo === "Fleet") {
            isFleetFullyEvaluated(valor, true).then(async ok => {
              if (ok) {
                const lastEval = await findLatestEvaluacionForFleet(valor, true);
                if (lastEval) return showExistingEvaluacionNotice(lastEval);
              }
              const existente = await checkEvaluacionExistenteTarget(tipo, valor);
              if (existente) showExistingEvaluacionNotice(existente);
              else clearExistingEvaluacionNotice();
            });
          } else if (tipo === "ProcesoSistema") {
            isProcesoFullyEvaluated(valor, state.wizard.fleet, true).then(async ok => {
              if (ok) {
                const lastEval = await findLatestEvaluacionForProceso(valor, state.wizard.fleet, true);
                if (lastEval) return showExistingEvaluacionNotice(lastEval);
              }
              const existente = await checkEvaluacionExistenteTarget(tipo, valor);
              if (existente) showExistingEvaluacionNotice(existente);
              else clearExistingEvaluacionNotice();
            });
          } else if (tipo === "EGI") {
            isEgiFullyEvaluated(valor, state.wizard.fleet, state.wizard.proceso, true).then(async ok => {
              if (ok) {
                const lastEval = await findLatestEvaluacionForEgi(valor, state.wizard.fleet, state.wizard.proceso, true);
                if (lastEval) return showExistingEvaluacionNotice(lastEval);
              }
              const existente = await checkEvaluacionExistenteTarget(tipo, valor);
              if (existente) showExistingEvaluacionNotice(existente);
              else clearExistingEvaluacionNotice();
            });
          }
        }
        });
      });

      on("txt-equipo", "input", (e) => searchEquipoLocal(e.target.value));
      on("txt-equipo", "focus", (e) => searchEquipoLocal(e.target.value));
    let searchTimer = null;
      on("escenario-search", "input", (e) => {
      state.wizard.escenarioSearch = e.target.value;
      state.wizard.escenarioPage = 0;
      clearTimeout(searchTimer);
      const term = state.wizard.escenarioSearch.trim();
      if (term.length >= 3 && state.wizard.escenarioThrottled) {
        searchTimer = setTimeout(async () => {
          try {
            showError("");
            const escenarios = await loadScenarioCandidatesBySearch(term);
            state.wizard.escenarioCandidates = escenarios;
            state.wizard.escenarioPage = 0;
            renderScenarioCandidates(escenarios);
            if (!escenarios.length) {
              showError("No hay escenarios para ese código. Intente con otro término.");
            }
          } catch (err) {
            showError("Error al buscar escenarios: " + err.message);
          }
        }, 300);
      } else {
        renderScenarioCandidates(state.wizard.escenarioCandidates);
      }
      });
      on("escenario-prev", "click", () => {
      if (state.wizard.escenarioPage > 0) {
        state.wizard.escenarioPage--;
        renderScenarioCandidates(state.wizard.escenarioCandidates);
      }
      });
      on("escenario-next", "click", () => {
      const filtered = filterScenarioCandidates(state.wizard.escenarioCandidates);
      if ((state.wizard.escenarioPage + 1) * state.wizard.escenarioPageSize < filtered.length) {
        state.wizard.escenarioPage++;
        renderScenarioCandidates(state.wizard.escenarioCandidates);
      }
      });
      on("escenario-page", "change", (e) => {
      const filtered = filterScenarioCandidates(state.wizard.escenarioCandidates);
      const totalPages = Math.max(1, Math.ceil(filtered.length / state.wizard.escenarioPageSize));
      let page = Number(e.target.value || 1);
      if (Number.isNaN(page) || page < 1) page = 1;
      if (page > totalPages) page = totalPages;
      state.wizard.escenarioPage = page - 1;
      renderScenarioCandidates(state.wizard.escenarioCandidates);
      });
      on("equipo-prev", "click", () => { if (state.search.page > 0) { state.search.page--; renderEquipoResults(); } });
      on("equipo-next", "click", () => { if ((state.search.page + 1) * state.search.pageSize < state.search.results.length) { state.search.page++; renderEquipoResults(); } });
      on("equipo-page", "change", (e) => {
      const totalPages = Math.max(1, Math.ceil(state.search.results.length / state.search.pageSize));
      let page = Number(e.target.value || 1);
      if (Number.isNaN(page) || page < 1) page = 1;
      if (page > totalPages) page = totalPages;
      state.search.page = page - 1;
      renderEquipoResults();
      });

      on("to-step-3", "click", async () => {
      showError("");
      setWizardState();
      if (state.wizard.nivel === "Fleet") {
        if (await isFleetFullyEvaluated(state.wizard.fleet, true)) {
          const lastEval = await findLatestEvaluacionForFleet(state.wizard.fleet, true);
          if (lastEval) showExistingEvaluacionNotice(lastEval);
          return;
        }
      }
      if (state.wizard.nivel === "ProcesoSistema") {
        if (await isProcesoFullyEvaluated(state.wizard.proceso, state.wizard.fleet, true)) {
          const lastEval = await findLatestEvaluacionForProceso(state.wizard.proceso, state.wizard.fleet, true);
          if (lastEval) showExistingEvaluacionNotice(lastEval);
          return;
        }
      }
      if (state.wizard.nivel === "EGI") {
        if (await isEgiFullyEvaluated(state.wizard.egi, state.wizard.fleet, state.wizard.proceso, true)) {
          const lastEval = await findLatestEvaluacionForEgi(state.wizard.egi, state.wizard.fleet, state.wizard.proceso, true);
          if (lastEval) showExistingEvaluacionNotice(lastEval);
          return;
        }
      }
      state.wizard.coverage = computeCoverageLists();
      state.wizard.equiposCubiertos = state.wizard.coverage.equipos.length;
      el("equipo-count-target").textContent = formatCoverageSummary(state.wizard.coverage);
      if (state.wizard.escenarioThrottled) {
        showError("La lista superó el umbral. Escriba al menos 3 caracteres en 'Buscar escenario' para filtrar por código.");
        setStep(3);
        return;
      }
      try {
        const escenarios = await loadScenarioCandidates();
        state.wizard.escenarioCandidates = escenarios;
        state.wizard.escenario = null;
        state.wizard.escenarioSearch = "";
        state.wizard.escenarioThrottled = false;
        state.wizard.escenarioPage = 0;
        el("escenario-search").value = "";
        el("evento-analizado").value = "";
        el("to-step-4").disabled = true;
        showError("");
        if (!escenarios.length) {
          showError("No existen escenarios activos para esta selección. Contacta al administrador.");
          return;
        }
        renderScenarioCandidates(escenarios);
        setStep(3);
      } catch (err) {
        if (String(err.message || "").includes("SPQueryThrottledException")) {
          state.wizard.escenarioThrottled = true;
          showError("La lista superó el umbral. Escriba al menos 3 caracteres en 'Buscar escenario' para filtrar por código o pide crear índices en EscenariosAC.");
        } else {
          showError("Error al cargar escenarios: " + err.message);
        }
      }
      });

      on("back-step-2", "click", () => {
      state.wizard.escenario = null;
      el("escenario-meta").textContent = "-";
      el("evento-analizado").value = "";
      state.wizard.eventoAnalizado = "";
      const edit = el("evento-analizado-edit");
      if (edit) edit.value = "";
      el("to-step-4").disabled = true;
      state.wizard.escenarioPage = 0;
      setStep(2);
      });

      on("to-step-4", "click", () => {
      const edit = el("evento-analizado-edit");
      if (edit) edit.value = state.wizard.eventoAnalizado || "";
      renderCoverageSummaryStep4();
      setStep(4);
      });
      on("back-step-3", "click", () => setStep(3));

      on("to-step-5", "click", () => {
      const respuestas = {};
      let ok = true;
      document.querySelectorAll("#preguntas select").forEach(select => {
        respuestas[select.dataset.key] = select.value;
        if (!select.value) ok = false;
      });
      const justificacion = el("justificacion").value.trim();
      if (!ok || !justificacion) {
        swalInfo("Faltan datos", "Complete todas las respuestas y la justificación.");
        return;
      }
      state.wizard.respuestas = respuestas;
      state.wizard.justificacion = justificacion;
      buildResumen();
      setStep(5);
      });

    const eventoEdit = el("evento-analizado-edit");
    if (eventoEdit) {
      eventoEdit.addEventListener("input", () => {
        state.wizard.eventoAnalizado = eventoEdit.value;
      });
    }

      on("back-step-4", "click", () => setStep(4));

      on("confirmar", "click", async () => {
        const confirmar = el("confirmar");
        if (confirmar) confirmar.disabled = true;
      try {
        await guardarEvaluacion();
        setStep(6);
        loadDashboard();
        loadMisEvaluaciones();
        await swalSuccess("Evaluación guardada", "La evaluación fue registrada correctamente.");
      } catch (err) {
        swalError("Error al guardar", err.message);
      } finally {
          if (confirmar) confirmar.disabled = false;
      }
      });

      on("nueva-otra", "click", () => resetWizard());
    }

    const adminNuevo = el("admin-nuevo");
    if (adminNuevo) adminNuevo.addEventListener("click", () => openAdminEditor(null));
    const adminCancelar = el("admin-cancelar");
    if (adminCancelar) adminCancelar.addEventListener("click", () => el("admin-editor").classList.add("hidden"));
    const adminGuardar = el("admin-guardar");
    if (adminGuardar) adminGuardar.addEventListener("click", saveAdminEscenario);

    on("admin-eval-refresh", "click", loadAdminEvaluaciones);
    const reqRefresh = el("admin-req-refresh");
    if (reqRefresh) reqRefresh.addEventListener("click", loadEditRequests);
    ["admin-eval-term", "admin-eval-desde", "admin-eval-hasta"].forEach(id => {
      const node = el(id);
      if (node) node.addEventListener("change", loadAdminEvaluaciones);
    });
    on("admin-eval-prev", "click", () => {
      if (state.adminEval.page > 0) {
        state.adminEval.page -= 1;
        renderAdminEvaluaciones();
      }
    });
    on("admin-eval-next", "click", () => {
      const totalPages = Math.max(1, Math.ceil((state.adminEval.items || []).length / state.adminEval.pageSize));
      if (state.adminEval.page + 1 < totalPages) {
        state.adminEval.page += 1;
        renderAdminEvaluaciones();
      }
    });
    on("admin-eval-page", "change", (e) => {
      const totalPages = Math.max(1, Math.ceil((state.adminEval.items || []).length / state.adminEval.pageSize));
      let page = Number(e.target.value || 1);
      if (!Number.isFinite(page)) page = 1;
      page = Math.min(Math.max(1, page), totalPages);
      state.adminEval.page = page - 1;
      renderAdminEvaluaciones();
    });
    on("dash-eval-prev", "click", () => {
      if (state.dashEval.page > 0) {
        state.dashEval.page -= 1;
        renderDashEvaluaciones();
      }
    });
    on("dash-eval-next", "click", () => {
      const totalPages = Math.max(1, Math.ceil((state.dashEval.items || []).length / state.dashEval.pageSize));
      if (state.dashEval.page + 1 < totalPages) {
        state.dashEval.page += 1;
        renderDashEvaluaciones();
      }
    });
    on("dash-eval-page", "change", (e) => {
      const totalPages = Math.max(1, Math.ceil((state.dashEval.items || []).length / state.dashEval.pageSize));
      let page = Number(e.target.value || 1);
      if (!Number.isFinite(page)) page = 1;
      page = Math.min(Math.max(1, page), totalPages);
      state.dashEval.page = page - 1;
      renderDashEvaluaciones();
    });
    on("mis-eval-prev", "click", () => {
      if (state.misEval.page > 0) {
        state.misEval.page -= 1;
        renderMisEvaluaciones();
      }
    });
    on("mis-eval-next", "click", () => {
      const totalPages = Math.max(1, Math.ceil((state.misEval.items || []).length / state.misEval.pageSize));
      if (state.misEval.page + 1 < totalPages) {
        state.misEval.page += 1;
        renderMisEvaluaciones();
      }
    });
    on("mis-eval-page", "change", (e) => {
      const totalPages = Math.max(1, Math.ceil((state.misEval.items || []).length / state.misEval.pageSize));
      let page = Number(e.target.value || 1);
      if (!Number.isFinite(page)) page = 1;
      page = Math.min(Math.max(1, page), totalPages);
      state.misEval.page = page - 1;
      renderMisEvaluaciones();
    });

    const allPrev = el("all-eval-prev");
    if (allPrev) allPrev.addEventListener("click", () => {
      if (state.allEval.page > 0) {
        state.allEval.page -= 1;
        renderAllEvaluaciones();
      }
    });
    const allNext = el("all-eval-next");
    if (allNext) allNext.addEventListener("click", () => {
      const totalPages = Math.max(1, Math.ceil((state.allEval.items || []).length / state.allEval.pageSize));
      if (state.allEval.page + 1 < totalPages) {
        state.allEval.page += 1;
        renderAllEvaluaciones();
      }
    });
    const allPage = el("all-eval-page");
    if (allPage) allPage.addEventListener("change", (e) => {
      const totalPages = Math.max(1, Math.ceil((state.allEval.items || []).length / state.allEval.pageSize));
      let page = Number(e.target.value || 1);
      if (!Number.isFinite(page)) page = 1;
      page = Math.min(Math.max(1, page), totalPages);
      state.allEval.page = page - 1;
      renderAllEvaluaciones();
    });

    ["admin-nivel", "admin-fleet", "admin-proceso", "admin-egi"].forEach(id => {
      const sel = el(id);
      if (sel) sel.addEventListener("change", loadAdminEscenarios);
    });
  }

  async function init() {
    state.user = await spHttp.getCurrentUser();
    setUserProfile(state.user);
    state.isAdmin = isAdmin(state.user.Groups || []);
    await loadUserRoles();
    if (state.role.canAdmin) state.isAdmin = true;
    const navAdmin = el("nav-admin");
    if (navAdmin) navAdmin.classList.toggle("hidden", !state.isAdmin);

    fillSelect(el("admin-nivel"), ["Corporativo", "Fleet", "ProcesoSistema", "EGI", "Equipo"], "Todos");
    fillSelect(el("sc-nivel"), ["Corporativo", "Fleet", "ProcesoSistema", "EGI", "Equipo"], "Seleccione...");

    buildPreguntas();
    buildEvalEditPreguntas();
    wireEvents();
    await loadEquipos();
    await loadDashboard();
    if (el("view-mis")) await loadMisEvaluaciones();
    if (el("view-all")) await loadAllEvaluaciones();
    await loadAdvancedDashboard();
    if (state.isAdmin && el("view-admin")) {
      await loadAdminEvaluaciones();
      await loadEditRequests();
    }
  }

  return { init };
})();

document.addEventListener("DOMContentLoaded", () => {
  app.init().catch(err => alert("Error inicializando: " + err.message));
});
