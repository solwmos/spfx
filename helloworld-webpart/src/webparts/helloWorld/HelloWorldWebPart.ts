import { Version,Environment,EnvironmentType}from '@microsoft/sp-core-library';
import {IPropertyPaneConfiguration,PropertyPaneTextField,PropertyPaneCheckbox,PropertyPaneDropdown} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import MockHttpClient from './MockHttpClient';
import { SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';

export interface IHelloWorldWebPartProps 
{
  description: string;
  MultiText: string;
  Show: boolean;
  color: string;
}
export interface ISPLists 
{
  value: ISPList[];
}
export interface ISPList 
{
  Title: string;
  Description: string;
}
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> 
{
  public render(): void 
  {
    var listSection =``;
    var introDiv = `
      <div class="${ styles.helloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to ${(this.context.pageContext.web.title)}</span>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <ul>
                <li>Absolute Url : ${(this.context.pageContext.web.absoluteUrl)}</li>
                <li>Template : ${(this.context.pageContext.web.templateName)}</li>
              </ul>
            </div>
          </div>
        </div>
      </div>`;
    if (this.properties.Show) 
    {
      listSection =`
      <div class="${ styles.helloWorld}" style="margin-top:5%;">
        <div class="${ styles.container }">
          <div class="${styles.row}">
            <div class="${ styles.column }">
              <h1 class="${styles.title}">Available Sharepoint Lists</h1>
              <span class="${ styles.subTitle }">Find below all the available sharepoint lists of the site.</span>
              <p class="${ styles.description }"></p>
              <div id="spListContainer" />
            </div>
          </div>
        </div>
      </div>`;
      this._renderListAsync();
    }
    this.domElement.innerHTML = introDiv+listSection;
  }
  private _renderListAsync(): void 
  {
    // Local environment
    if (Environment.type === EnvironmentType.Local) 
    {
      this._getMockListData().then((response) => { this._renderList(response.value); });
    }
    else if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) 
    {
      this._getListData().then((response) => { this._renderList(response.value);
        });
    }
  }
  private _getMockListData(): Promise<ISPLists> 
  {
    return MockHttpClient.get().then((data: ISPList[]) => 
    {
        var listData: ISPLists = { value: data };
        return listData;
    }) as Promise<ISPLists>;
  }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration 
  {
    return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
            PropertyPaneTextField('description', {
              label: 'Description'
            }),
            PropertyPaneTextField('MultiText', {
              label: 'Please give a intro text',
              multiline: true
            }),
            PropertyPaneCheckbox('Show', {
              text: 'Show Section'
            }),
            PropertyPaneDropdown('Color', {
              label: 'Dropdown',
              options: [
                { key: 'row-themeDark', text: 'themeDark' },
                { key: 'row-themeDarkAlt', text: 'themeDarkAlt' },
                { key: 'row-themeDarker', text: 'themeDarker' }
              ]})
          ]
          }
        ]
      }
    ]
    };
  }
  private _getListData(): Promise<ISPLists> 
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {return response.json();});
  }
  private _renderList(items: ISPList[]): void 
  {
    let html: string = '';
    html = `<ul class="${styles.list}">`;
    items.forEach((item: ISPList) => {
      html += `
        <li class="${styles.listItem}">
          <h3><b>${item.Title}</b></h3>
          <span>${item.Description}</span>
        </li>`;
    });
    html += `</ul>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }
}
