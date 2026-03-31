var React = require("react");
var FluentUI = require("@fluentui/react");
var DetailsList = FluentUI.DetailsList;
var SelectionMode = FluentUI.SelectionMode;
var DetailsListLayoutMode = FluentUI.DetailsListLayoutMode;
var Checkbox = FluentUI.Checkbox;
var Dropdown = FluentUI.Dropdown;
var PrimaryButton = FluentUI.PrimaryButton;
var DefaultButton = FluentUI.DefaultButton;
var MessageBar = FluentUI.MessageBar;
var MessageBarType = FluentUI.MessageBarType;
var Spinner = FluentUI.Spinner;
var SpinnerSize = FluentUI.SpinnerSize;
var Icon = FluentUI.Icon;
var ProgressIndicator = FluentUI.ProgressIndicator;
var Stack = FluentUI.Stack;
var FluentText = FluentUI.Text;

var GroupingType = require("../models/GroupingType");
var GS = GroupingType.GroupingStrategy;
var MaxLevels = GroupingType.MaxLevelsForStrategy;
var SharePointService = require("../services/SharePointService").SharePointService;
var styles = require("./ListFolderOrganizer.module.scss");

/* ------------------------------------------------------------------ */
/*  Dropdown option helpers                                            */
/* ------------------------------------------------------------------ */

var strategyOptions = [
  { key: GS.Date, text: "By Date" },
  { key: GS.Text, text: "By Initial Letter" },
  { key: GS.Choice, text: "By Choice Value" }
];

function buildLevelsOptions(strategy) {
  var max = MaxLevels[strategy] || 1;
  var opts = [];
  for (var i = 1; i <= max; i++) {
    var label = i + " level" + (i > 1 ? "s" : "");
    if (strategy === GS.Date) {
      if (i === 1) label += " (Year)";
      if (i === 2) label += " (Year / Month)";
      if (i === 3) label += " (Year / Month / Day)";
    } else if (strategy === GS.Text) {
      if (i === 1) label += " (1st letter)";
      if (i === 2) label += " (1st + 2nd letter)";
      if (i === 3) label += " (1st + 2nd + 3rd letter)";
    }
    opts.push({ key: i, text: label });
  }
  return opts;
}

/* ================================================================== */
/*  Component                                                          */
/* ================================================================== */

function ListFolderOrganizer(props) {
  var listsState = React.useState([]);
  var lists = listsState[0];
  var setLists = listsState[1];

  var loadingState = React.useState(true);
  var loading = loadingState[0];
  var setLoading = loadingState[1];

  var errorState = React.useState("");
  var error = errorState[0];
  var setError = errorState[1];

  var processingState = React.useState(false);
  var processing = processingState[0];
  var setProcessing = processingState[1];

  var logState = React.useState([]);
  var logMessages = logState[0];
  var setLogMessages = logState[1];

  var service = React.useMemo(function () {
    return new SharePointService(props.spHttpClient, props.siteUrl);
  }, [props.spHttpClient, props.siteUrl]);

  /* ----- initial load ----- */
  React.useEffect(function () {
    loadLists();
  }, []);

  function loadLists() {
    setLoading(true);
    setError("");
    service.getLists()
      .then(function (result) {
        setLists(result);
        // Fetch root item counts in the background
        result.forEach(function (list) {
          service.getRootItemCount(list.id, list.rootFolderUrl)
            .then(function (count) {
              setLists(function (prev) {
                return prev.map(function (l) {
                  return l.id === list.id ? Object.assign({}, l, { rootItemCount: count }) : l;
                });
              });
            })
            .catch(function () { /* swallow */ });
        });
      })
      .catch(function (err) {
        setError("Failed to load lists: " + (err.message || err));
      })
      .then(function () {
        setLoading(false);
      });
  }

  /* ----- row-level handlers ----- */

  function onSelectToggle(listId, checked) {
    setLists(function (prev) {
      return prev.map(function (l) {
        return l.id === listId ? Object.assign({}, l, { selected: checked }) : l;
      });
    });

    if (checked) {
      var list = lists.filter(function (l) { return l.id === listId; })[0];
      if (list && list.availableFields.length === 0) {
        service.getListFields(listId, list.groupingStrategy)
          .then(function (fields) {
            setLists(function (prev) {
              return prev.map(function (l) {
                return l.id === listId ? Object.assign({}, l, { availableFields: fields }) : l;
              });
            });
          })
          .catch(function (err) {
            setError("Failed to load fields: " + (err.message || err));
          });
      }
    }
  }

  function onStrategyChange(listId, strategy) {
    setLists(function (prev) {
      return prev.map(function (l) {
        return l.id === listId
          ? Object.assign({}, l, {
              groupingStrategy: strategy,
              sourceFieldInternalName: "",
              sourceFieldTitle: "",
              availableFields: [],
              levels: 1
            })
          : l;
      });
    });

    service.getListFields(listId, strategy)
      .then(function (fields) {
        setLists(function (prev) {
          return prev.map(function (l) {
            return l.id === listId ? Object.assign({}, l, { availableFields: fields }) : l;
          });
        });
      })
      .catch(function (err) {
        setError("Failed to load fields: " + (err.message || err));
      });
  }

  function onLevelsChange(listId, levels) {
    setLists(function (prev) {
      return prev.map(function (l) {
        return l.id === listId ? Object.assign({}, l, { levels: levels }) : l;
      });
    });
  }

  function onSourceFieldChange(listId, fieldInternalName, fieldTitle) {
    setLists(function (prev) {
      return prev.map(function (l) {
        return l.id === listId
          ? Object.assign({}, l, { sourceFieldInternalName: fieldInternalName, sourceFieldTitle: fieldTitle })
          : l;
      });
    });
  }

  /* ----- logging helpers ----- */
  function addLog(msg) {
    setLogMessages(function (prev) {
      return prev.concat(["[" + new Date().toLocaleTimeString() + "] " + msg]);
    });
  }

  function updateListStatus(listId, status, message, progress) {
    setLists(function (prev) {
      return prev.map(function (l) {
        return l.id === listId
          ? Object.assign({}, l, { status: status, statusMessage: message, progress: progress })
          : l;
      });
    });
  }

  /* ----- main action ----- */

  function organizeSelected() {
    var selected = lists.filter(function (l) {
      return l.selected && l.sourceFieldInternalName;
    });
    if (selected.length === 0) {
      setError("Select at least one list and configure its grouping (strategy + source field).");
      return;
    }

    setProcessing(true);
    setLogMessages([]);
    setError("");

    var promise = Promise.resolve();

    selected.forEach(function (list) {
      promise = promise.then(function () {
        return processOneList(list);
      });
    });

    promise.then(function () {
      setProcessing(false);
      addLog("— All done! —");
    });
  }

  function processOneList(list) {
    updateListStatus(list.id, "Processing", "Starting…", 0);
    addLog("Processing list: " + list.title);

    var p = Promise.resolve();

    /* 1. Enable folders if needed */
    if (!list.folderCreationEnabled) {
      p = p.then(function () {
        addLog('  Enabling folder creation for "' + list.title + '"…');
        return service.enableFolderCreation(list.id).then(function () {
          setLists(function (prev) {
            return prev.map(function (l) {
              return l.id === list.id ? Object.assign({}, l, { folderCreationEnabled: true }) : l;
            });
          });
          addLog("  ✓ Folder creation enabled");
        });
      });
    }

    /* 2. Get all items */
    var items;
    p = p.then(function () {
      updateListStatus(list.id, "Processing", "Fetching items…", 5);
      addLog("  Fetching items…");
      return service.getListItems(list.id, list.sourceFieldInternalName)
        .then(function (result) {
          items = result;
          addLog("  Found " + items.length + " items");
        });
    });

    /* 3-5. Calculate paths, create folders, move items */
    p = p.then(function () {
      if (items.length === 0) {
        updateListStatus(list.id, "Done", "No items to organize.", 100);
        addLog("  No items found — skipping.");
        return;
      }

      var folderPaths = [];
      var seen = {};
      var itemMoves = [];

      for (var idx = 0; idx < items.length; idx++) {
        var item = items[idx];
        var folderPath = SharePointService.generateFolderPath(
          item, list.sourceFieldInternalName, list.groupingStrategy, list.levels
        );
        if (!seen[folderPath]) {
          seen[folderPath] = true;
          folderPaths.push(folderPath);
        }
        var targetDir = list.rootFolderUrl + "/" + folderPath;
        if (item.FileDirRef !== targetDir) {
          itemMoves.push({ itemId: item.Id, folderPath: folderPath });
        }
      }

      /* 4. Create folders */
      addLog("  Creating " + folderPaths.length + " folder(s)…");
      var folderPromise = Promise.resolve();
      folderPaths.forEach(function (fp, fi) {
        folderPromise = folderPromise.then(function () {
          return service.ensureFolder(list.rootFolderUrl, fp).then(function () {
            updateListStatus(
              list.id, "Processing",
              "Creating folders (" + (fi + 1) + "/" + folderPaths.length + ")…",
              5 + Math.round(((fi + 1) / folderPaths.length) * 25)
            );
          });
        });
      });

      return folderPromise.then(function () {
        addLog("  ✓ " + folderPaths.length + " folder(s) ensured");

        if (itemMoves.length === 0) {
          addLog("  All items already in correct folders.");
          updateListStatus(list.id, "Done", "All items already organized.", 100);
          return;
        }

        /* 5. Move items */
        addLog("  Moving " + itemMoves.length + " item(s) to folders…");
        var movePromise = Promise.resolve();
        var moveCount = 0;
        itemMoves.forEach(function (move) {
          movePromise = movePromise.then(function () {
            var dest = list.rootFolderUrl + "/" + move.folderPath;
            return service.moveItemToFolder(list.id, move.itemId, dest)
              .catch(function (moveErr) {
                addLog("  ⚠ Failed to move item " + move.itemId + ": " + (moveErr.message || moveErr));
              })
              .then(function () {
                moveCount++;
                if (moveCount % 10 === 0 || moveCount === itemMoves.length) {
                  updateListStatus(
                    list.id, "Processing",
                    "Moving items (" + moveCount + "/" + itemMoves.length + ")…",
                    30 + Math.round((moveCount / itemMoves.length) * 70)
                  );
                }
              });
          });
        });

        return movePromise.then(function () {
          addLog("  ✓ " + moveCount + " item(s) organized");
          updateListStatus(list.id, "Done", "Done! " + moveCount + " items organized.", 100);
        });
      });
    });

    return p.catch(function (err) {
      addLog("  ✗ Error: " + (err.message || err));
      updateListStatus(list.id, "Error", err.message || String(err), 0);
    });
  }

  /* ----- Select / Deselect All ----- */
  function selectAll() {
    setLists(function (prev) {
      return prev.map(function (l) { return Object.assign({}, l, { selected: true }); });
    });
    lists.forEach(function (list) {
      if (list.availableFields.length === 0) {
        service.getListFields(list.id, list.groupingStrategy)
          .then(function (fields) {
            setLists(function (prev) {
              return prev.map(function (l) {
                return l.id === list.id ? Object.assign({}, l, { availableFields: fields }) : l;
              });
            });
          })
          .catch(function () { /* swallow */ });
      }
    });
  }

  function deselectAll() {
    setLists(function (prev) {
      return prev.map(function (l) { return Object.assign({}, l, { selected: false }); });
    });
  }

  /* ================================================================ */
  /*  Column definitions                                               */
  /* ================================================================ */

  var columns = [
    {
      key: "selected",
      name: "",
      minWidth: 32,
      maxWidth: 32,
      onRender: function (item) {
        return React.createElement(Checkbox, {
          checked: item.selected,
          onChange: function (_ev, checked) { onSelectToggle(item.id, !!checked); },
          disabled: processing
        });
      }
    },
    {
      key: "title",
      name: "List Name",
      fieldName: "title",
      minWidth: 140,
      maxWidth: 250,
      isResizable: true
    },
    {
      key: "itemCount",
      name: "Total",
      fieldName: "itemCount",
      minWidth: 50,
      maxWidth: 60
    },
    {
      key: "rootItemCount",
      name: "In Root",
      minWidth: 60,
      maxWidth: 75,
      onRender: function (item) {
        if (item.rootItemCount === -1) {
          return React.createElement(FluentText, { title: "List exceeds 5 000-item view threshold" }, "> 5,000");
        }
        return React.createElement(FluentText, null, item.rootItemCount);
      }
    },
    {
      key: "folderEnabled",
      name: "Folders",
      minWidth: 55,
      maxWidth: 55,
      onRender: function (item) {
        return React.createElement(Icon, {
          iconName: item.folderCreationEnabled ? "CheckMark" : "Cancel",
          className: item.folderCreationEnabled ? styles.folderEnabled : styles.folderDisabled,
          title: item.folderCreationEnabled
            ? "Folder creation enabled"
            : "Folder creation disabled – will be auto-enabled"
        });
      }
    },
    {
      key: "strategy",
      name: "Grouping",
      minWidth: 140,
      maxWidth: 170,
      onRender: function (item) {
        return React.createElement(Dropdown, {
          selectedKey: item.groupingStrategy,
          options: strategyOptions,
          onChange: function (_ev, option) {
            if (option) onStrategyChange(item.id, option.key);
          },
          disabled: processing,
          className: styles.cellDropdown
        });
      }
    },
    {
      key: "levels",
      name: "Depth",
      minWidth: 165,
      maxWidth: 210,
      onRender: function (item) {
        return React.createElement(Dropdown, {
          selectedKey: item.levels,
          options: buildLevelsOptions(item.groupingStrategy),
          onChange: function (_ev, option) {
            if (option) onLevelsChange(item.id, option.key);
          },
          disabled: processing,
          className: styles.cellDropdown
        });
      }
    },
    {
      key: "sourceField",
      name: "Source Field",
      minWidth: 160,
      maxWidth: 220,
      onRender: function (item) {
        return React.createElement(Dropdown, {
          selectedKey: item.sourceFieldInternalName || undefined,
          options: item.availableFields.map(function (f) {
            return { key: f.internalName, text: f.title + " (" + f.typeAsString + ")" };
          }),
          placeholder: item.availableFields.length === 0
            ? "Check row to load fields"
            : "Select field…",
          onChange: function (_ev, option) {
            if (option) {
              var field = item.availableFields.filter(function (f) {
                return f.internalName === option.key;
              })[0];
              onSourceFieldChange(item.id, option.key, field ? field.title : "");
            }
          },
          disabled: processing || item.availableFields.length === 0,
          className: styles.cellDropdown
        });
      }
    },
    {
      key: "status",
      name: "Status",
      minWidth: 180,
      maxWidth: 280,
      onRender: function (item) {
        if (item.status === "Processing") {
          return React.createElement(Stack, null,
            React.createElement(ProgressIndicator, {
              percentComplete: item.progress / 100,
              description: item.statusMessage
            })
          );
        }
        if (item.status === "Done") {
          return React.createElement(FluentText, { className: styles.statusDone },
            React.createElement(Icon, { iconName: "CheckMark" }), " ", item.statusMessage
          );
        }
        if (item.status === "Error") {
          return React.createElement(FluentText, { className: styles.statusError },
            React.createElement(Icon, { iconName: "ErrorBadge" }), " ", item.statusMessage
          );
        }
        return React.createElement(FluentText, { className: styles.statusIdle }, "Ready");
      }
    }
  ];

  /* ================================================================ */
  /*  Render                                                           */
  /* ================================================================ */

  var content;
  if (loading) {
    content = React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading lists…" });
  } else if (lists.length === 0) {
    content = React.createElement(MessageBar, { messageBarType: MessageBarType.info },
      "No custom lists found on this site."
    );
  } else {
    content = React.createElement("div", { className: styles.tableContainer },
      React.createElement(DetailsList, {
        items: lists,
        columns: columns,
        selectionMode: SelectionMode.none,
        layoutMode: DetailsListLayoutMode.justified,
        isHeaderVisible: true,
        compact: false
      })
    );
  }

  var errorBar = null;
  if (error) {
    errorBar = React.createElement(MessageBar, {
      messageBarType: MessageBarType.error,
      onDismiss: function () { setError(""); },
      isMultiline: false
    }, error);
  }

  var logPanel = null;
  if (logMessages.length > 0) {
    logPanel = React.createElement("div", { className: styles.logPanel },
      React.createElement(FluentText, {
        variant: "smallPlus",
        style: { fontWeight: 600, marginBottom: 4, display: "block" }
      }, "Activity Log"),
      logMessages.map(function (msg, i) {
        return React.createElement("div", { key: i, className: styles.logLine }, msg);
      })
    );
  }

  return React.createElement("div", { className: styles.listFolderOrganizer },
    React.createElement(Stack, { tokens: { childrenGap: 16 } },
      /* Header */
      React.createElement(Stack, {
        horizontal: true,
        horizontalAlign: "space-between",
        verticalAlign: "center",
        wrap: true,
        tokens: { childrenGap: 8 }
      },
        React.createElement(FluentText, { variant: "xLarge", className: styles.title },
          React.createElement(Icon, { iconName: "FolderList" }), " List Folder Organizer"
        ),
        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 8 } },
          React.createElement(DefaultButton, {
            text: "Select All", onClick: selectAll, disabled: processing || loading
          }),
          React.createElement(DefaultButton, {
            text: "Deselect All", onClick: deselectAll, disabled: processing || loading
          }),
          React.createElement(DefaultButton, {
            iconProps: { iconName: "Refresh" }, text: "Refresh",
            onClick: loadLists, disabled: processing
          }),
          React.createElement(PrimaryButton, {
            iconProps: { iconName: "Play" }, text: "Organize Selected",
            onClick: organizeSelected, disabled: processing || loading
          })
        )
      ),
      errorBar,
      content,
      logPanel
    )
  );
}

module.exports = ListFolderOrganizer;
module.exports.default = ListFolderOrganizer;
