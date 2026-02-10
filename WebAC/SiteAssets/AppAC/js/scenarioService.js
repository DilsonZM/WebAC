const scenarioService = (() => {
  const listName = "EscenariosAC";
  const esc = (v) => String(v || "").replace(/'/g, "''");
  let cachedListId = null;

  async function getListId() {
    if (cachedListId) return cachedListId;
    const url = `${spHttp.getWebUrl()}/_api/web/lists/GetByTitle('${listName.replace(/'/g, "''")}')?$select=Id`;
    const data = await spHttp.get(url);
    cachedListId = data.Id;
    return cachedListId;
  }

  async function listCandidates(targetTipo, targetValor, top = 200, searchTerm = "") {
    const term = String(searchTerm || "").trim();
    let filter = "";
    if (term.length >= 3) {
      filter = `startswith(Title,'${esc(term)}')`;
    } else {
      if (targetTipo === "Fleet") filter = `Fleet eq '${esc(targetValor)}'`;
      if (targetTipo === "ProcesoSistema") filter = `ProcesoSistema eq '${esc(targetValor)}'`;
      if (targetTipo === "EGI") filter = `EGI eq '${esc(targetValor)}'`;
      if (targetTipo === "Equipo") filter = `EquipNo eq '${esc(targetValor)}'`;
    }
    // Activo may be unset in some items; filter client-side instead

    const select = "Id,Title,EventoAnalizado,EquipNo,Fleet,ProcesoSistema,EGI,Activo";
    const url = spHttp.listItemsUrl(listName, `$select=${select}&$filter=${encodeURIComponent(filter)}&$orderby=Id&$top=${top}`);
    const data = await spHttp.get(url);
    return (data.value || []).filter(item => item.Activo !== false);
  }

  async function searchCandidates(term, top = 50) {
    const listId = await getListId();
    const safeTerm = esc(term);
    const querytext = `ListId:${listId} AND Title:${safeTerm}*`;
    const selectprops = "Title,ListItemID,Path";
    const url = `${spHttp.getWebUrl()}/_api/search/query?querytext='${encodeURIComponent(querytext)}'&trimduplicates=true&rowlimit=${top}&selectproperties='${encodeURIComponent(selectprops)}'`;
    const data = await spHttp.search(url);
    const rows = data?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    return rows.map(r => {
      const cells = r.Cells || [];
      const getVal = (key) => (cells.find(c => c.Key === key) || {}).Value;
      return {
        Title: getVal("Title"),
        ListItemID: Number(getVal("ListItemID"))
      };
    }).filter(x => Number.isFinite(x.ListItemID));
  }

  async function getById(id) {
    const select = "Id,Title,EventoAnalizado,EquipNo,Fleet,ProcesoSistema,EGI,Activo";
    const url = spHttp.listItemsUrl(listName, `$select=${select}&$filter=Id eq ${id}`);
    const data = await spHttp.get(url);
    return (data.value || [])[0] || null;
  }

  return { listCandidates, searchCandidates, getById };
})();
