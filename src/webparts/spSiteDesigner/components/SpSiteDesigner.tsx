import * as React from 'react';
import styles from './SpSiteDesigner.module.scss';
import { ISpSiteDesignerProps } from './ISpSiteDesignerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';
export interface ISpSiteDesignerState {
  siteScriptResults: any;
  siteScriptData: string;
}

export default class SpSiteDesigner extends React.Component<ISpSiteDesignerProps, ISpSiteDesignerState> {
  constructor(props: any) {
    super(props);
    this.state = {
      siteScriptData: null,
      siteScriptResults: null
    }
    this._handleSiteScriptChange = this._handleSiteScriptChange.bind(this);
  }
  // private RestRequest(url: string, params?: any): any {
  //   const req = new XMLHttpRequest();
  //   req.onreadystatechange = function () {
  //     if (req.readyState != 4) // Loaded
  //       return;
  //     console.log(req.responseText);
  //   };

  //   // Prepend web URL to url and remove duplicated slashes.
  //   const webBasedUrl = (this.props.context.pageContext.web.absoluteUrl + url).replace(/\/{2,}/, "/");
  //   console.log(this.props.context.pageContext.web.absoluteUrl);
  //   console.log(webBasedUrl);

  //   // https://itgroovedeveloper.sharepoint.com/itgroovedeveloper.sharepoint.com/sites/mattdev///_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title=%27Contoso%20theme%20script%27
  //   console.log(url)

  //   req.open("POST", webBasedUrl, true);
  //   req.setRequestHeader("Content-Type", "application/json;charset=utf-8");
  //   req.setRequestHeader("ACCEPT", "application/json; odata.metadata=minimal");
  //   req.setRequestHeader("x-requestdigest", this.props.context.pageContext.web.absoluteUrl.requestdigest);
  //   req.setRequestHeader("ODATA-VERSION", "4.0");
  //   req.send(params ? JSON.stringify(params) : void 0);
  // }

  

  private site_script: string =
  `{
    "$schema": "schema.json",
    "actions": [
      {
        "verb": "applyTheme",
        "themeName": "Contoso Theme"
      }
    ],
    "bindata": {},
    "version": 1
  }`;

  private _createSiteScript(siteScriptData: any): any {

    const spOpts: ISPHttpClientOptions = {
      body: siteScriptData
    };

    const url = "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title='Contoso theme script'";
    return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}${url}`, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: JSON) => {
          console.log(responseJSON);
        });
      });
  }

  private _getSiteScripts(): any {
    const url = "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts";
    return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}${url}`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          this.setState({
            siteScriptResults: responseJSON.value
          })
        });
      });
  }

  // private _getSiteScriptMetadata(id: string): any {
  //   return this.RestRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata",
  //     { id: id });
  // }

  private _getSiteScriptMetadata(id: string): any {

    const spOpts: ISPHttpClientOptions = {
      body: JSON.stringify(id),
      headers: {
        'Content-Type': 'application/json;charset=utf-8',
        ACCEPT: 'application/json; odata.metadata=minimal',
        'ODATA-VERSION': '4.0'
      }
    };

    const url = "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata";
    return this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}${url}`, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: JSON) => {
          console.log(responseJSON);
        });
      });
  }

  // private _createSiteDesign(): any {
  //   return this.RestRequest("/_api / Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign", {
  //     info: {
  //       Title: "Contoso customer tracking",
  //       Description: "Creates customer list and applies standard theme",
  //       SiteScriptIds: ["07702c07-0485-426f-b710-4704241caad9"],
  //       WebTemplate: "64",
  //       PreviewImageUrl: "https://contoso.sharepoint.com/SiteAssets/contoso-design.png",
  //       PreviewImageAltText: "Customer tracking site design theme"
  //     }
  //   });
  // }

  // private _getSiteDesign(): any {
  //   return this.RestRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns");
  // }

  // private _getSiteDesignMetadata(): any {
  //   this.RestRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata",
  //     { id: "614f9b28-3e85-4ec9-a961-5971ea086cca" });
  // }

  private _handleCreateSiteScriptClick(): any {
    this._createSiteScript(this.state.siteScriptData);
  }

  private _handleGetSiteScriptClick(): any {
    console.log(this._getSiteScripts());
  }

  private _handleSiteScriptChange(event: any): any {
    this.setState({
      siteScriptData: event.target.value
    });
  }

  private _handleSiteScriptEdit(id: string): any {
    const siteScript: any = {
      id: id
    }
    this._getSiteScriptMetadata(siteScript);
  }

  public render(): React.ReactElement<ISpSiteDesignerProps> {
    const { siteScriptResults } = this.state;
    return (
      <div className={styles.spSiteDesigner}>
        <textarea id="siteScript" value={this.state.siteScriptData} onChange={this._handleSiteScriptChange}></textarea>
        <button onClick={() => this._handleCreateSiteScriptClick()}>Create Site Script</button>

        <button onClick={() => this._handleGetSiteScriptClick()}>Get Site Scripts</button>

        <ul>
        {siteScriptResults && siteScriptResults.map(siteScript => 
            <li>{siteScript.Title} 
              <button onClick={() => this._handleSiteScriptEdit(siteScript.Id)}>Edit</button>
            </li>
          )}
        </ul>

      </div>
    );
  }
}
