// Grouping strategy constants (plain JS - no TypeScript enums)
var GroupingStrategy = {
  Date: "Date",
  Text: "Text",
  Choice: "Choice"
};

var GroupingStrategyLabels = {
  "Date": "By Date",
  "Text": "By Initial Letter",
  "Choice": "By Choice Value"
};

var MaxLevelsForStrategy = {
  "Date": 3,
  "Text": 3,
  "Choice": 1
};

module.exports = {
  GroupingStrategy: GroupingStrategy,
  GroupingStrategyLabels: GroupingStrategyLabels,
  MaxLevelsForStrategy: MaxLevelsForStrategy
};
