import * as React from 'react';
import styles from './SpSiteDesigner.module.scss';
import { ISpSiteDesignerProps } from './ISpSiteDesignerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';
export interface ISpSiteDesignerState {
  siteScriptResults?: any;
  siteDesignResults?: any;
  siteScriptData?: string;
  siteScriptTitle?: string;
  siteDesignTitle?: string;
  siteDesignDescription?: string;
  siteDesignWebTemplate?: number;
  siteDesignPreviewImageUrl?: string;
  siteDesignPreviewImageAltText?: string;
}

export default class SpSiteDesigner extends React.Component<ISpSiteDesignerProps, ISpSiteDesignerState> {
  constructor(props: any) {
    super(props);
    this.state = {
      siteScriptData: null,
      siteScriptResults: null,
      siteScriptTitle: null
    }
    this._handleInputChange = this._handleInputChange.bind(this);
  }


  public baseUrl: string = '/';

  private _getEffectiveUrl(relativeUrl: string): string {
    return (this.baseUrl + '//' + relativeUrl).replace(/\/{2,}/, '/');
  }

  private _restRequest(url: string, params: any = null): Promise<any> {
    const restUrl = this._getEffectiveUrl(url);
    const options: ISPHttpClientOptions = {
      body: JSON.stringify(params),
      headers: {
        'Content-Type': 'application/json;charset=utf-8',
        ACCEPT: 'application/json; odata.metadata=minimal',
        'ODATA-VERSION': '4.0'
      }
    };
    return this.props.context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, options).then((response) => {
      if (response.status == 204) {
        return {};
      } else {
        return response.json();
      }
    });
  }

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

  private _createSiteScript(siteScriptTitle: string, siteScriptData: any): any {
    siteScriptData = JSON.parse(siteScriptData);
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title='${siteScriptTitle}'`,
      siteScriptData
    );
  }

  private _getSiteScripts(): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`
    ).then((response) => {
      this.setState({
        siteScriptResults: response.value
      })
    });
  }

  private _getSiteScriptMetadata(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata`,
      id
    ).then((response) => {
      console.log(response);
    });
  }

  private _deleteSiteScript(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript`,
      id
    ).then((response) => {
      console.log(response);
    });
  }

  private _createSiteDesign(): any {
    return this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign", {
      info: {
        Title: this.state.siteDesignTitle,
        Description: this.state.siteDesignDescription,
        SiteScriptIds: ["07702c07-0485-426f-b710-4704241caad9"],
        WebTemplate: this.state.siteDesignWebTemplate,
        PreviewImageUrl: this.state.siteDesignPreviewImageUrl,
        PreviewImageAltText: this.state.siteDesignPreviewImageAltText
      }
    }).then((response) => {
      console.log(response);
    });
  }

  private _getSiteDesigns(): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`
    ).then((response) => {
      this.setState({
        siteDesignResults: response.value
      })
    });
  }

  // private _getSiteDesignMetadata(): any {
  //   this.RestRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata",
  //     { id: "614f9b28-3e85-4ec9-a961-5971ea086cca" });
  // }

  private _handleCreateSiteScriptClick(): any {
    this._createSiteScript(this.state.siteScriptTitle, this.state.siteScriptData);
  }

  private _handleCreateSiteDesignClick(): any {
    this._createSiteDesign();
  }

  private _handleGetSiteScriptClick(): any {
    this._getSiteScripts();
  }

  private _handleGetSiteDesignClick(): any {
    this._getSiteDesigns();
  }

  private _handleInputChange(event: any): any {
    const target = event.target;
    const value = target.type === 'checkbox' ? target.checked : target.value;
    const name = target.name;
    console.log(name)
    this.setState({
      [name]: value
    });
  }

  private _handleSiteScriptEdit(id: string): any {
    const siteScript: any = {
      id: id
    }
    this._getSiteScriptMetadata(siteScript);
  }

  private _handleDeleteSiteScript(id: string): any {
    const siteScript: any = {
      id: id
    }
    let shouldDelete: boolean = confirm("Are you sure you want to delete this script?");
    if (shouldDelete) {
      this._deleteSiteScript(siteScript);
    }
  }

  public _groupSiteScriptsBySiteDesign(siteScripts, siteDesigns) {
    const siteDesignsWithSiteScripts = [];
    // each site design
    for (var i = 0; i < siteDesigns.length; i++) {
      const siteScriptsInSiteDesign = [];
      // each script in the design
      for (var k = 0; k < siteDesigns[i].SiteScriptIds.length; k++) {
        // compare to overall list of scripts
        for (var j = 0; j < siteScripts.length; j++) {
          if (siteDesigns[i].SiteScriptIds[k] === siteScripts[j].Id) {
            siteScriptsInSiteDesign.push(siteScripts[j]);
          }
        }
      }
      siteDesigns[i].SiteScripts = siteScriptsInSiteDesign;
      siteDesignsWithSiteScripts.push(siteDesigns[i]);
    }
    return siteDesignsWithSiteScripts;
  }

  public render(): React.ReactElement<ISpSiteDesignerProps> {
    const { siteScriptResults, siteDesignResults } = this.state;

    let siteDesignsWithSiteScripts;
    if (siteScriptResults && siteScriptResults) {
      siteDesignsWithSiteScripts = this._groupSiteScriptsBySiteDesign(siteScriptResults, siteDesignResults);
    }

    return (
      <div className={styles.spSiteDesigner} >

        <button onClick={() => this._handleGetSiteScriptClick()}>Get Site Scripts</button>
        <button onClick={() => this._handleGetSiteDesignClick()}>Get Site Designs</button>

        <div>
          <form>
            <div><input id="siteScriptTitle" name="siteScriptTitle" value={this.state.siteScriptTitle} onChange={this._handleInputChange}></input></div>
            <div><textarea id="siteScriptData" name="siteScriptData" value={this.state.siteScriptData} onChange={this._handleInputChange}></textarea></div>
          </form>
          <button onClick={() => this._handleCreateSiteScriptClick()}>Create Site Script</button>
        </div>



        <div>
          <form>
            <div><input name="siteDesignTitle" value={this.state.siteDesignTitle} onChange={this._handleInputChange} /></div>
            <div><input name="siteDesignDescription" value={this.state.siteDesignDescription} onChange={this._handleInputChange} /></div>
            <div><input name="siteDesignWebTemplate" value={this.state.siteDesignWebTemplate} onChange={this._handleInputChange} /></div>
            <div><input name="siteDesignPreviewImageUrl" value={this.state.siteDesignPreviewImageUrl} onChange={this._handleInputChange} /></div>
            <div><input name="siteDesignPreviewImageAltText" value={this.state.siteDesignPreviewImageAltText} onChange={this._handleInputChange} /></div >
          </form>
          <button onClick={() => this._handleCreateSiteDesignClick()}>Create Site Design</button>
        </div>

        <ul>
          {siteScriptResults && siteScriptResults.map(siteScript =>
            <li>{siteScript.Title}
              <button onClick={() => this._handleSiteScriptEdit(siteScript.Id)}>Edit</button>
              <button onClick={() => this._handleDeleteSiteScript(siteScript.Id)}>Delete</button>
            </li>
          )}
        </ul>

        <ul>
          {siteDesignResults && siteDesignResults.map(siteDesign =>
            <li>{siteDesign.Title}
              {/* <button onClick={() => this._handlesiteDesignEdit(siteDesign.Id)}>Edit</button>
              <button onClick={() => this._handleDeletesiteDesign(siteDesign.Id)}>Delete</button> */}
            </li>
          )}
        </ul>

        {siteDesignsWithSiteScripts && 
          <ul>
            {siteDesignsWithSiteScripts.map(siteDesign =>
            <li>{siteDesign.Title}</li>
            )}
          </ul>
        
        }
        

      </div>
    );
  }
}
