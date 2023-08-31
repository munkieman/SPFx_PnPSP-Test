import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnPspTestWebPart.module.scss';
import * as strings from 'PnPspTestWebPartStrings';
//import {
//  SPHttpClient,
//  SPHttpClientResponse
//} from '@microsoft/sp-http';

import {spfi, SPFx} from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { LogLevel, PnPLogging } from "@pnp/logging";
//import { Caching } from "@pnp/queryable";

export interface IPnPspTestWebPartProps {
  description: string;
}

export default class PnPspTestWebPart extends BaseClientSideWebPart<IPnPspTestWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  //private LOG_SOURCE = "🅿PnPjsExample";

  public getData():void{
    let html : string = "";
    let folderName : string = "";
    let folderNamePrev : string = "";

    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));

    sp.web.lists.getByTitle('Procedures').items
      .select("*","FieldValuesAsText/MetaInfo","TaxCatchAll/Term","FieldValuesAsText/FileLeafRef")  //"FieldValuesAsText/FileLeafRef", 
      //.filter("Team eq 'ASM Team B'")
      .expand("FieldValuesAsText","TaxCatchAll/Term")
      .getAll()
      .then((results: any[])=>{
        if(results.length>0){
          console.log(results);

          let x=0;
          results.forEach(()=>{            
            folderName = results[x].TaxCatchAll[0].Term;
            if(folderName !== folderNamePrev){
              html+=`<li>${results[x].TaxCatchAll[0].Term}</li>`;
              console.log(results[x].TaxCatchAll[0].Term);
              folderNamePrev=folderName;
            }
            x++;
          }); 
          
          const listContainer: Element = this.domElement.querySelector('#terms');
          listContainer.innerHTML = html;
                
          //console.log(html); 
          //this.properties.Title = results.Title;
        }
      });
  }

  public render(): void {

    //const graph = graphfi().using(graphSPFx(this.context));
    //const spCache = spfi(sp).using(Caching({store:"session"}));
    //const data = spCache.web.lists.getByTitle("Countries").items;
    
    this.domElement.innerHTML = `
    <section class="${styles.pnPspTest} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>        
      </div>
      Data</br>
      <ul id="terms"> 
      </ul>
    </section>`;

    this.getData();
  }

  protected async Init(): Promise<void> {
    await super.onInit();
   
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
