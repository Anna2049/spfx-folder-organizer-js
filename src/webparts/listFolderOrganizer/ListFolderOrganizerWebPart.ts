var React = require("react");
var ReactDom = require("react-dom");
var sp_webpart_base = require("@microsoft/sp-webpart-base");
var sp_property_pane = require("@microsoft/sp-property-pane");
var _ListFolderOrganizer = require("./components/ListFolderOrganizer").default;

var ListFolderOrganizerWebPart = /** @class */ (function (_super) {
  // Manually set up prototype chain (ES5 inheritance)
  function ListFolderOrganizerWebPart() {
    return _super !== null && _super.apply(this, arguments) || this;
  }

  // Inherit from BaseClientSideWebPart
  ListFolderOrganizerWebPart.prototype = Object.create(_super.prototype);
  ListFolderOrganizerWebPart.prototype.constructor = ListFolderOrganizerWebPart;

  ListFolderOrganizerWebPart.prototype.render = function () {
    var element = React.createElement(_ListFolderOrganizer, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl
    });
    ReactDom.render(element, this.domElement);
  };

  ListFolderOrganizerWebPart.prototype.onDispose = function () {
    ReactDom.unmountComponentAtNode(this.domElement);
  };

  ListFolderOrganizerWebPart.prototype.getPropertyPaneConfiguration = function () {
    return {
      pages: [
        {
          header: { description: "List Folder Organizer Settings" },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                sp_property_pane.PropertyPaneTextField("description", {
                  label: "Description"
                })
              ]
            }
          ]
        }
      ]
    };
  };

  return ListFolderOrganizerWebPart;
}(sp_webpart_base.BaseClientSideWebPart));

module.exports = { default: ListFolderOrganizerWebPart };
