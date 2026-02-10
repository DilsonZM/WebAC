const catalogService = (() => {
  const listName = "EquiposAC";
  const normalize = (value) => String(value || "").trim().toUpperCase();

  function mapSearchRow(row) {
    const cells = (row.Cells || row.cells || []);
    const out = {};
    cells.forEach(c => {
      if (!c || !c.Key) return;
      out[String(c.Key).toLowerCase()] = c.Value;
    });
    return out;
  }

  function cleanOwsValue(value) {
    if (value === null || value === undefined) return "";
    const text = String(value);
    if (!text.includes(";#")) return text;
    const parts = text.split(";#").filter(Boolean);
    return parts.length ? parts[parts.length - 1] : "";
  }

  function pick(row, ...keys) {
    for (const key of keys) {
      const v = row[String(key).toLowerCase()];
      if (v !== undefined && v !== null && v !== "") return v;
    }
    return "";
  }

  async function getListMeta() {
    const webUrl = spHttp.getWebUrl();
    const safeName = listName.replace(/'/g, "''");
    const url = `${webUrl}/_api/web/lists/GetByTitle('${safeName}')?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    return spHttp.get(url);
  }

  async function loadEquiposById(select, batchSize = 2000) {
    const items = [];
    let lastId = 0;
    while (true) {
      const query = `$select=${select}&$filter=Id gt ${lastId}&$orderby=Id&$top=${batchSize}`;
      const url = spHttp.listItemsUrl(listName, query);
      const data = await spHttp.get(url);
      const batch = data?.value || [];
      if (!batch.length) break;
      items.push(...batch);
      lastId = Math.max(...batch.map(i => i.Id || 0));
      if (batch.length < batchSize) break;
    }
    return items;
  }

  async function loadEquiposSearch() {
    const webUrl = spHttp.getWebUrl();
    const meta = await getListMeta();
    const listId = meta?.Id || meta?.id || "";
    const select = [
      "Title",
      "ows_Title",
      "NombreEquipo",
      "ows_NombreEquipo",
      "Fleet",
      "ows_Fleet",
      "ProcesoSistema",
      "ows_ProcesoSistema",
      "EGI",
      "ows_EGI",
      "BranchGerencia",
      "ows_BranchGerencia",
      "SiteSuperintendencia",
      "ows_SiteSuperintendencia",
      "UnidadProceso",
      "ows_UnidadProceso",
      "Estado",
      "ows_Estado",
      "ListItemID",
      "ows_ListItemID"
    ].join(",");
    const rowLimit = 500;
    let start = 0;
    let all = [];
    while (true) {
      const queryText = listId ? `ListId:${listId}` : `contentclass:STS_ListItem`;
      const query = `querytext='${queryText}'&selectproperties='${select}'&rowlimit=${rowLimit}&startrow=${start}`;
      const url = `${webUrl}/_api/search/query?${query}`;
      const data = await spHttp.search(url);
      const rows = data?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
      if (!rows.length) break;
      const items = rows.map(mapSearchRow).map(r => ({
        Id: Number(cleanOwsValue(pick(r, "ListItemID", "ows_ListItemID"))) || 0,
        Title: cleanOwsValue(pick(r, "Title", "ows_Title")),
        NombreEquipo: cleanOwsValue(pick(r, "NombreEquipo", "ows_NombreEquipo")),
        Fleet: cleanOwsValue(pick(r, "Fleet", "ows_Fleet")),
        ProcesoSistema: cleanOwsValue(pick(r, "ProcesoSistema", "ows_ProcesoSistema")),
        EGI: cleanOwsValue(pick(r, "EGI", "ows_EGI")),
        BranchGerencia: cleanOwsValue(pick(r, "BranchGerencia", "ows_BranchGerencia")),
        SiteSuperintendencia: cleanOwsValue(pick(r, "SiteSuperintendencia", "ows_SiteSuperintendencia")),
        UnidadProceso: cleanOwsValue(pick(r, "UnidadProceso", "ows_UnidadProceso")),
        Estado: cleanOwsValue(pick(r, "Estado", "ows_Estado"))
      }));
      all = all.concat(items);
      if (rows.length < rowLimit) break;
      start += rowLimit;
    }
    return all;
  }

  async function loadEquipos(pageSize) {
    const select = "Id,Title,NombreEquipo,Fleet,ProcesoSistema,EGI,BranchGerencia,SiteSuperintendencia,UnidadProceso,Estado";
    let items = [];
    try {
      items = await loadEquiposById(select, Math.max(pageSize || 0, 2000));
    } catch (err) {
      items = await loadEquiposSearch();
    }
    const equiposMap = new Map();
    items.forEach(i => {
      const titleKey = normalize(i.Title);
      if (titleKey) equiposMap.set(titleKey, i);
      if (i.Title) equiposMap.set(i.Title, i);
    });
    return {
      equipos: items,
      equiposMap,
      distinct: {
        fleet: [...new Set(items.map(i => normalize(i.Fleet)).filter(Boolean))].sort(),
        proceso: [...new Set(items.map(i => normalize(i.ProcesoSistema)).filter(Boolean))].sort(),
        egi: [...new Set(items.map(i => normalize(i.EGI)).filter(Boolean))].sort(),
        gerencia: [...new Set(items.map(i => normalize(i.BranchGerencia)).filter(Boolean))].sort(),
        superintendencia: [...new Set(items.map(i => normalize(i.SiteSuperintendencia)).filter(Boolean))].sort(),
        unidadProceso: [...new Set(items.map(i => normalize(i.UnidadProceso)).filter(Boolean))].sort()
      }
    };
  }

  return { loadEquipos };
})();
