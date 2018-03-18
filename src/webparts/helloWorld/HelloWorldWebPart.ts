import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  IWebPartContext
} from '@microsoft/sp-webpart-base';
import { Version } from '@microsoft/sp-core-library';
import { IHelloWorldWebPartProps } from './IHelloWorldWebPartProps';
import * as $ from 'jquery';
import * as moment from 'moment';
require('./FeedEK.js');


export default class RssReaderWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private guid: string;

  public constructor(context?: IWebPartContext) {
    super();

    this.guid = this.getGuid();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public render(): void {

    var html = `
    <div id="RssTitle" style="font-family: Segoe UI; font-weight: lighter; font-size: ${this.properties.titleSize}px;">
      ${this.properties.title == undefined ? "Latest News" : this.properties.title}
    </div><div id="${this.guid}"></div>`;
    html += `
<style>
<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.4.0/css/fabric.min.css">
.feedEkList{width:100%; list-style:none outside none; border:0px solid #D3CAD7; padding:4px 6px; color:#3E3E3E;}
.feedEkList li{border-bottom:1px solid #D3CAD7; padding:5px;}
.feedEkList li:last-child{border-bottom:none;}
.itemTitle a{font-weight:bold; font-family:"Segoe UI"; text-decoration:none }
.itemTitle a:hover{ text-decoration:underline }
.itemDate{font-size:11px;color:#AAAAAA;}
</style>
    `;
    this.domElement.innerHTML = html;

        ($ as any)('#' + this.guid).FeedEk({
            FeedUrl: this.properties.feedUrl,
            MaxCount : this.properties.maxCount,
            ShowDesc : this.properties.showDesc,
            ShowPubDate: this.properties.showPubDate,
            DescCharacterLimit: this.properties.descCharacterLimit,
            DateFormat: this.properties.dateFormat,
            DateFormatLang: this.properties.dateFormatLang
        });
  }

  /**
   * @function
   * Generates a GUID
   */
  private getGuid(): string {
    return this.s4() + this.s4() + '-' + this.s4() + '-' + this.s4() + '-' +
      this.s4() + '-' + this.s4() + this.s4() + this.s4();
  }

  /**
   * @function
   * Generates a GUID part
   */
  private s4(): string {
      return Math.floor((1 + Math.random()) * 0x10000)
        .toString(16)
        .substring(1);
  }


  /**
   * @function
   * PropertyPanel settings definition
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "RSS Reader Properties:"
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Preferences",
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Title Text",
                  value: "Latest News"
                }),
                PropertyPaneTextField('titleSize', {
                  label: "Title Size (In pixels)"
                }),
                PropertyPaneTextField('feedUrl', {
                  label: "Feed URL"
                }),
                PropertyPaneSlider('maxCount', {
                  label: "Max Count",
                  min: 1,
                  max: 100,
                  step: 1
                }),
                PropertyPaneToggle('showPubDate', {
                  label: "Show Publishing Date"
                }),
                PropertyPaneToggle('showDesc', {
                  label: "Show Description"
                }),
                PropertyPaneSlider('descCharacterLimit', {
                  label: "Character Limit in Descriptions",
                  min: 1,
                  max: 500,
                  step: 1
                })
              ]
            }
          ]
        }
      ]
    };
  }
}