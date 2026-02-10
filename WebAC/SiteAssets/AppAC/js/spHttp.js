const spHttp = (() => {
  const getWebUrl = () => {
    if (window._spPageContextInfo?.webAbsoluteUrl) {
      return window._spPageContextInfo.webAbsoluteUrl;
    }
    const parts = location.pathname.split("/").filter(Boolean);
    if (parts[0] === "sites" && parts.length >= 3) {
      return `${location.origin}/sites/${parts[1]}/${parts[2]}`;
    }
    return location.origin;
  };

  const listItemTypeCache = new Map();
  const listFieldMapCache = new Map();
  const listFieldMapCacheAll = new Map();

  async function getRequestDigest(webUrl) {
    const res = await fetch(`${webUrl}/_api/contextinfo`, { method: "POST", headers: { "Accept": "application/json;odata=verbose" } });
    const data = await res.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  async function get(url) {
    const res = await fetch(url, { headers: { "Accept": "application/json;odata=nometadata" } });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
  }

  async function search(queryUrl) {
    const res = await fetch(queryUrl, { headers: { "Accept": "application/json;odata=verbose" } });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
  }

  async function getListItemEntityType(listName) {
    if (listItemTypeCache.has(listName)) return listItemTypeCache.get(listName);
    const safeName = listName.replace(/'/g, "''");
    const url = `${getWebUrl()}/_api/web/lists/GetByTitle('${safeName}')?$select=ListItemEntityTypeFullName`;
    const data = await get(url);
    const typeName = data.ListItemEntityTypeFullName;
    listItemTypeCache.set(listName, typeName);
    return typeName;
  }

  async function getListFields(listName) {
    if (listFieldMapCache.has(listName)) return listFieldMapCache.get(listName);
    const safeName = listName.replace(/'/g, "''");
    const url = `${getWebUrl()}/_api/web/lists/GetByTitle('${safeName}')/fields?$select=InternalName,TypeAsString,Hidden,ReadOnlyField`;
    const data = await get(url);
    const map = {};
    (data.value || []).forEach(f => {
      if (f.Hidden || f.ReadOnlyField) return;
      map[f.InternalName] = f.TypeAsString;
    });
    listFieldMapCache.set(listName, map);
    return map;
  }

  async function getListFieldsAll(listName) {
    if (listFieldMapCacheAll.has(listName)) return listFieldMapCacheAll.get(listName);
    const safeName = listName.replace(/'/g, "''");
    const url = `${getWebUrl()}/_api/web/lists/GetByTitle('${safeName}')/fields?$select=InternalName,TypeAsString,Hidden,ReadOnlyField`;
    const data = await get(url);
    const map = {};
    (data.value || []).forEach(f => {
      if (f.Hidden) return;
      map[f.InternalName] = f.TypeAsString;
    });
    listFieldMapCacheAll.set(listName, map);
    return map;
  }

  function coercePayloadByFieldTypes(payload, fieldMap) {
    const out = {};
    Object.keys(payload || {}).forEach(key => {
      if (key === "__metadata") { out[key] = payload[key]; return; }
      const isIdKey = key.endsWith("Id");
      const baseKey = isIdKey ? key.slice(0, -2) : key;
      const type = fieldMap[baseKey] || fieldMap[key];
      if (!type) return;
      const isLookupType = type === "User" || type === "Lookup" || type === "UserMulti" || type === "LookupMulti";
      if (isLookupType && !isIdKey) return;
      if (!isLookupType && isIdKey) return;
      const val = payload[key];
      if (val === undefined) return;
      if (val === null) { out[key] = null; return; }

      switch (type) {
        case "Text":
        case "Note":
        case "Choice":
        case "MultiChoice":
        case "URL":
        case "Computed":
          out[key] = String(val);
          break;
        case "Number":
        case "Currency":
        case "Integer":
          out[key] = val === "" ? null : Number(val);
          break;
        case "Boolean":
          out[key] = Boolean(val);
          break;
        case "DateTime":
          out[key] = typeof val === "string" ? val : new Date(val).toISOString();
          break;
        case "User":
        case "Lookup":
          out[key] = Number(val);
          break;
        case "UserMulti":
        case "LookupMulti":
          if (val === null) { out[key] = null; break; }
          if (val && typeof val === "object" && Array.isArray(val.results)) {
            out[key] = { results: val.results.map(Number) };
          } else if (Array.isArray(val)) {
            out[key] = { results: val.map(Number) };
          } else {
            out[key] = { results: [Number(val)] };
          }
          break;
        default:
          out[key] = val;
          break;
      }
    });
    return out;
  }

  async function post(url, body) {
    const webUrl = getWebUrl();
    const digest = await getRequestDigest(webUrl);
    const res = await fetch(url, {
      method: "POST",
      headers: { "Accept": "application/json;odata=nometadata", "Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest },
      body: JSON.stringify(body)
    });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
  }

  async function postListItem(listName, body) {
    const typeName = await getListItemEntityType(listName);
    const payload = body && body.__metadata ? body : { __metadata: { type: typeName }, ...body };
    const webUrl = getWebUrl();
    const digest = await getRequestDigest(webUrl);
    const url = listItemsUrl(listName);
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      },
      body: JSON.stringify(payload)
    });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
  }

  async function merge(url, body) {
    const webUrl = getWebUrl();
    const digest = await getRequestDigest(webUrl);
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*"
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) throw new Error(await res.text());
    return null;
  }

  async function fetchPaged(url) {
    const items = [];
    const seen = new Set();
    let next = url;
    while (next) {
      const data = await get(next);
      const value = data.value || [];
      value.forEach(item => {
        const key = item.Id ? `${item.Id}` : JSON.stringify(item);
        if (!seen.has(key)) {
          seen.add(key);
          items.push(item);
        }
      });
      next = data["@odata.nextLink"] || null;
    }
    return items;
  }

  function listItemsUrl(listName, query) {
    const safeName = listName.replace(/'/g, "''");
    const base = `${getWebUrl()}/_api/web/lists/GetByTitle('${safeName}')/items`;
    return query ? `${base}?${query}` : base;
  }

  async function getCurrentUser() {
    const url = `${getWebUrl()}/_api/web/currentuser?$expand=Groups`;
    return get(url);
  }

  return { getWebUrl, getRequestDigest, get, post, postListItem, merge, fetchPaged, listItemsUrl, getCurrentUser, search, getListItemEntityType, getListFields, getListFieldsAll, coercePayloadByFieldTypes };
})();
