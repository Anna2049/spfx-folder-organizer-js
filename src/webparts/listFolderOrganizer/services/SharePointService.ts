var SPHttpClient = require("@microsoft/sp-http").SPHttpClient;
var _GS = require("../models/GroupingType").GroupingStrategy;

/* ------------------------------------------------------------------ */
/*  SharePointService - all REST API operations (pure JavaScript)      */
/* ------------------------------------------------------------------ */

function SharePointService(spHttpClient, siteUrl) {
  this._spHttpClient = spHttpClient;
  this._siteUrl = siteUrl;
}

/* ------------------------------------------------------------------ */
/*  Lists                                                              */
/* ------------------------------------------------------------------ */

SharePointService.prototype.getLists = function () {
  var self = this;
  var apiUrl =
    self._siteUrl + "/_api/web/lists?" +
    "$filter=Hidden eq false and IsCatalog eq false and BaseTemplate eq 100" +
    "&$select=Id,Title,ItemCount,EnableFolderCreation,RootFolder/ServerRelativeUrl" +
    "&$expand=RootFolder" +
    "&$orderby=Title";

  return self._spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
    .then(function (response) {
      if (!response.ok) {
        return response.text().then(function (errText) {
          throw new Error("Failed to fetch lists: " + response.statusText + " — " + errText);
        });
      }
      return response.json();
    })
    .then(function (data) {
      return (data.value || []).map(function (list) {
        return {
          id: list.Id,
          title: list.Title,
          itemCount: list.ItemCount,
          rootItemCount: 0,
          folderCreationEnabled: list.EnableFolderCreation,
          rootFolderUrl: list.RootFolder.ServerRelativeUrl,
          selected: false,
          groupingStrategy: _GS.Date,
          levels: 1,
          sourceFieldInternalName: "",
          sourceFieldTitle: "",
          availableFields: [],
          status: "Idle",
          statusMessage: "",
          progress: 0
        };
      });
    });
};

SharePointService.prototype.getRootItemCount = function (listId, rootFolderUrl) {
  var self = this;
  var escaped = SharePointService._escapeODataString(rootFolderUrl);
  var apiUrl =
    self._siteUrl + "/_api/web/lists(guid'" + listId + "')/items?" +
    "$filter=FSObjType eq 0 and FileDirRef eq '" + escaped + "'" +
    "&$select=Id" +
    "&$top=5001";

  return self._spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
    .then(function (response) {
      if (!response.ok) {
        return -1;
      }
      return response.json().then(function (data) {
        var items = data.value || [];
        return items.length > 5000 ? -1 : items.length;
      });
    })
    .catch(function () {
      return -1;
    });
};

/* ------------------------------------------------------------------ */
/*  Fields                                                             */
/* ------------------------------------------------------------------ */

SharePointService.prototype.getListFields = function (listId, strategy) {
  var self = this;
  var typeFilter;
  if (strategy === _GS.Date) {
    typeFilter = "TypeAsString eq 'DateTime'";
  } else if (strategy === _GS.Text) {
    typeFilter = "(TypeAsString eq 'Text' or TypeAsString eq 'Note' or TypeAsString eq 'Choice')";
  } else if (strategy === _GS.Choice) {
    typeFilter = "(TypeAsString eq 'Choice' or TypeAsString eq 'MultiChoice')";
  } else {
    typeFilter = "TypeAsString eq 'Text'";
  }

  var apiUrl =
    self._siteUrl + "/_api/web/lists(guid'" + listId + "')/fields?" +
    "$filter=Hidden eq false and " + typeFilter +
    "&$select=Id,InternalName,Title,TypeAsString" +
    "&$orderby=Title";

  return self._spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
    .then(function (response) {
      if (!response.ok) {
        throw new Error("Failed to fetch fields: " + response.statusText);
      }
      return response.json();
    })
    .then(function (data) {
      return (data.value || []).map(function (field) {
        return {
          id: field.Id,
          internalName: field.InternalName,
          title: field.Title,
          typeAsString: field.TypeAsString
        };
      });
    });
};

/* ------------------------------------------------------------------ */
/*  Enable folders                                                     */
/* ------------------------------------------------------------------ */

SharePointService.prototype.enableFolderCreation = function (listId) {
  var self = this;
  var apiUrl = self._siteUrl + "/_api/web/lists(guid'" + listId + "')";
  var options = {
    headers: {
      "Accept": "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE"
    },
    body: JSON.stringify({ EnableFolderCreation: true })
  };

  return self._spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, options)
    .then(function (response) {
      if (!response.ok) {
        throw new Error("Failed to enable folder creation: " + response.statusText);
      }
    });
};

/* ------------------------------------------------------------------ */
/*  Folders                                                            */
/* ------------------------------------------------------------------ */

SharePointService.prototype.ensureFolder = function (listRootFolderUrl, relativeFolderPath) {
  var self = this;
  var parts = relativeFolderPath.split("/").filter(function (p) { return p.length > 0; });

  var promise = Promise.resolve();
  var currentPath = listRootFolderUrl;

  parts.forEach(function (part) {
    promise = promise.then(function () {
      currentPath = currentPath + "/" + part;
      var escaped = SharePointService._escapeODataString(currentPath);
      var checkUrl = self._siteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + escaped + "')";

      return self._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1)
        .then(function (checkResp) {
          if (checkResp.ok) {
            return checkResp.json().then(function (folderData) {
              if (folderData.Exists !== false) {
                return; // already there
              }
              return self._createFolder(escaped);
            });
          }
          return self._createFolder(escaped);
        });
    });
  });

  return promise;
};

SharePointService.prototype._createFolder = function (escapedPath) {
  var self = this;
  var createUrl = self._siteUrl + "/_api/web/folders/add('" + escapedPath + "')";
  return self._spHttpClient.post(
    createUrl,
    SPHttpClient.configurations.v1,
    { headers: { "Accept": "application/json;odata=nometadata" } }
  ).then(function (createResp) {
    if (!createResp.ok) {
      return createResp.text().then(function (errText) {
        if (errText.toLowerCase().indexOf("already exists") === -1) {
          throw new Error("Failed to create folder: " + errText);
        }
      });
    }
  });
};

/* ------------------------------------------------------------------ */
/*  Items                                                              */
/* ------------------------------------------------------------------ */

SharePointService.prototype.getListItems = function (listId, sourceFieldInternalName) {
  var self = this;
  var allItems = [];
  var apiUrl =
    self._siteUrl + "/_api/web/lists(guid'" + listId + "')/items?" +
    "$select=Id," + sourceFieldInternalName + ",FileDirRef,FileRef" +
    "&$filter=FSObjType eq 0" +
    "&$top=5000";

  function fetchPage(url) {
    return self._spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(function (response) {
        if (!response.ok) {
          throw new Error("Failed to fetch list items: " + response.statusText);
        }
        return response.json();
      })
      .then(function (data) {
        allItems = allItems.concat(data.value || []);
        var nextLink = data["odata.nextLink"] || data["@odata.nextLink"] || null;
        if (nextLink) {
          return fetchPage(nextLink);
        }
        return allItems;
      });
  }

  return fetchPage(apiUrl);
};

SharePointService.prototype.moveItemToFolder = function (listId, itemId, targetFolderServerRelativeUrl) {
  var self = this;
  var apiUrl =
    self._siteUrl + "/_api/web/lists(guid'" + listId + "')" +
    "/items(" + itemId + ")/ValidateUpdateListItem()";

  var options = {
    headers: {
      "Accept": "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata"
    },
    body: JSON.stringify({
      formValues: [
        {
          FieldName: "FileDirRef",
          FieldValue: targetFolderServerRelativeUrl
        }
      ],
      bNewDocumentUpdate: false
    })
  };

  return self._spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, options)
    .then(function (response) {
      if (!response.ok) {
        return response.text().then(function (errText) {
          throw new Error("Failed to move item " + itemId + ": " + response.statusText + " — " + errText);
        });
      }
      return response.json();
    })
    .then(function (result) {
      if (result.value) {
        for (var i = 0; i < result.value.length; i++) {
          if (result.value[i].HasException) {
            throw new Error("Error moving item " + itemId + ": " + result.value[i].ErrorMessage);
          }
        }
      }
    });
};

/* ------------------------------------------------------------------ */
/*  Folder-path generation (static helpers)                            */
/* ------------------------------------------------------------------ */

SharePointService.generateFolderPath = function (item, sourceFieldInternalName, strategy, levels) {
  var value = item[sourceFieldInternalName];
  if (value === null || value === undefined || value === "") {
    return "_Uncategorized";
  }

  if (strategy === _GS.Date) {
    return SharePointService._generateDatePath(value, levels);
  } else if (strategy === _GS.Text) {
    return SharePointService._generateTextPath(value, levels);
  } else if (strategy === _GS.Choice) {
    return SharePointService._generateChoicePath(value);
  }
  return "_Uncategorized";
};

SharePointService._generateDatePath = function (dateValue, levels) {
  var date = new Date(dateValue);
  if (isNaN(date.getTime())) return "_Uncategorized";

  var year = date.getFullYear().toString();
  var monthNum = SharePointService._pad2(date.getMonth() + 1);
  var monthName = date.toLocaleString("en-US", { month: "long" });
  var day = SharePointService._pad2(date.getDate());

  var parts = [year];
  if (levels >= 2) parts.push(monthNum + " - " + monthName);
  if (levels >= 3) parts.push(day);

  return parts.join("/");
};

SharePointService._generateTextPath = function (textValue, levels) {
  var str = String(textValue).trim();
  if (!str) return "_Uncategorized";

  var upper = str.toUpperCase();
  var firstChar = upper.charAt(0);
  var isLetter = /^[A-Z]$/.test(firstChar);
  var level1 = isLetter ? firstChar : "#";

  var parts = [level1];
  if (levels >= 2) {
    var chars2 = upper.substring(0, 2);
    parts.push(chars2.length >= 2 ? chars2 : level1);
  }
  if (levels >= 3) {
    var chars3 = upper.substring(0, 3);
    parts.push(chars3.length >= 3 ? chars3 : parts[parts.length - 1]);
  }

  return parts.join("/");
};

SharePointService._generateChoicePath = function (value) {
  if (!value) return "_Uncategorized";
  var parts = String(value).split(";#").filter(function (v) { return v.trim().length > 0; });
  var choiceValue = parts[0] || "_Uncategorized";
  return SharePointService._sanitizeFolderName(choiceValue);
};

/* ------------------------------------------------------------------ */
/*  Utility                                                            */
/* ------------------------------------------------------------------ */

SharePointService._sanitizeFolderName = function (name) {
  return name.replace(/[~#%&*{}\\:<>?\/+|"]/g, "_").trim() || "_Unnamed";
};

SharePointService._escapeODataString = function (value) {
  return value.replace(/'/g, "''");
};

SharePointService._pad2 = function (n) {
  return n < 10 ? "0" + n : "" + n;
};

module.exports = { SharePointService: SharePointService };
