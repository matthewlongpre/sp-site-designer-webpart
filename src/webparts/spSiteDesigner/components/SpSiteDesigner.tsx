import * as React from 'react';
import styles from './SpSiteDesigner.module.scss';
import { ISpSiteDesignerProps } from './ISpSiteDesignerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';

import DualListBox from 'react-dual-listbox';
import 'react-dual-listbox/lib/react-dual-listbox.css';

export interface ISpSiteDesignerState {
  siteScriptResults?: any;
  siteDesignResults?: any;
  siteScriptData?: string;
  siteScriptTitle?: string;
  siteDesignTitle?: string;
  selectedSiteDesignID?: string;
  siteDesignDescription?: string;
  siteDesignWebTemplate?: number;
  siteDesignPreviewImageUrl?: string;
  siteDesignPreviewImageAltText?: string;
  selectedSiteScripts?: any;
}

export default class SpSiteDesigner extends React.Component<ISpSiteDesignerProps, ISpSiteDesignerState> {
  constructor(props: any) {
    super(props);
    this.state = {
      siteScriptData: null,
      siteScriptResults: null,
      siteScriptTitle: null,
      selectedSiteScripts: []
    };
    this._handleInputChange = this._handleInputChange.bind(this);
  }

  public componentDidMount() {
    this._loadData();
  }

  public _loadData() {
    this._getSiteScripts();
    this._getSiteDesigns();
  }

  public _reset() {
    // this.setState({
    //   selectedSiteDesignID: undefined
    // });
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
      });
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
    if (this.state.selectedSiteDesignID) {
      return this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign", {
        updateInfo: {
          Title: this.state.siteDesignTitle,
          Id: this.state.selectedSiteDesignID,
          Description: this.state.siteDesignDescription,
          SiteScriptIds: this.state.selectedSiteScripts,
          WebTemplate: this.state.siteDesignWebTemplate,
          PreviewImageUrl: this.state.siteDesignPreviewImageUrl,
          PreviewImageAltText: this.state.siteDesignPreviewImageAltText
        }
      }).then((response) => {
        console.log(response);
        this._loadData();
      });
    }
    return this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign", {
      info: {
        Title: this.state.siteDesignTitle,
        Description: this.state.siteDesignDescription,
        SiteScriptIds: this.state.selectedSiteScripts,
        WebTemplate: this.state.siteDesignWebTemplate,
        PreviewImageUrl: this.state.siteDesignPreviewImageUrl,
        PreviewImageAltText: this.state.siteDesignPreviewImageAltText
      }
    }).then((response) => {
      console.log(response);
      this._loadData();
    });
  }

  private _deleteSiteDesign(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign`,
      id
    ).then((response) => {
      console.log(response);
      this._loadData();
    });
  }

  private _getSiteDesigns(): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`
    ).then((response) => {
      this.setState({
        siteDesignResults: response.value
      });
    });
  }

  private _getSiteDesignMetadata(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`,
      id
    ).then((response) => {
      console.log(response);
      this.setState({
        selectedSiteDesignID: response.Id,
        siteDesignTitle: response.Title,
        siteDesignDescription: response.Description,
        selectedSiteScripts: response.SiteScriptIds,
        siteDesignWebTemplate: response.WebTemplate,
        siteDesignPreviewImageUrl: response.PreviewImageUrl,
        siteDesignPreviewImageAltText: response.PreviewImageAltText
      });
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
    this._reset();
    this._getSiteScripts();
  }

  private _handleGetSiteDesignClick(): any {
    this._reset();
    this._getSiteDesigns();
  }

  private _handleInputChange(event: any): any {
    const target = event.target;
    const value = target.type === 'checkbox' ? target.checked : target.value;
    const name = target.name;
    this.setState({
      [name]: value
    });
  }

  private _handleSiteScriptEdit(id: string): any {
    const siteScript: any = {
      id: id
    };
    this._getSiteScriptMetadata(siteScript);
  }

  private _handleSiteDesignEdit(id: string): any {
    const siteDesign: any = {
      id: id
    };
    this._getSiteDesignMetadata(siteDesign);
  }

  private _handleDeleteSiteScript(id: string): any {
    const siteScript: any = {
      id: id
    };
    let shouldDelete: boolean = confirm("Are you sure you want to delete this script?");
    if (shouldDelete) {
      this._deleteSiteScript(siteScript);
    }
  }

  private _handleDeleteSiteDesign(id: string): any {
    const siteDesign: any = {
      id: id
    }
    let shouldDelete: boolean = confirm("Are you sure you want to delete this Design?");
    if (shouldDelete) {
      this._deleteSiteDesign(siteDesign);
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
    const { siteScriptResults, siteDesignResults, selectedSiteScripts } = this.state;

    let siteDesignsWithSiteScripts;
    if (siteScriptResults && siteScriptResults) {
      siteDesignsWithSiteScripts = this._groupSiteScriptsBySiteDesign(siteScriptResults, siteDesignResults);
    }

    let siteScriptOptions;
    if (siteScriptResults) {
      siteScriptOptions = siteScriptResults.map((option) => {
        let r: any = {};
        r.label = option.Title;
        r.value = option.Id;
        return r;
      });
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
            <div><input name="siteDesignPreviewImageAltText" value={this.state.siteDesignPreviewImageAltText} onChange={this._handleInputChange} /></div>

            <div style={{display: 'flex',justifyContent: 'space-between'}}>
              <div>Available Site Scripts</div>
              <div>Added to Site Design</div>
            </div>
            {siteScriptOptions && <DualListBox
              availableLabel="Available Site Scripts"
              selectedLabel="Added to Site Design"
              simpleValue={true}
              options={siteScriptOptions}
              selected={selectedSiteScripts}
              onChange={(selectedSiteScripts) => {
                this.setState({ selectedSiteScripts });
              }}
            />}
          </form>
          <button onClick={() => this._handleCreateSiteDesignClick()}>Create Site Design</button>
        </div>

        <ul>
          {siteDesignResults && siteDesignResults.map(siteDesign =>
            <li>{siteDesign.Title}
              <button onClick={() => this._handleSiteDesignEdit(siteDesign.Id)}>Edit</button>
              <button onClick={() => this._handleDeleteSiteDesign(siteDesign.Id)}>Delete</button>
            </li>
          )}
        </ul>

        {/* {siteDesignsWithSiteScripts &&
          <ul>
            {siteDesignsWithSiteScripts.map(siteDesign =>
              <li>{siteDesign.Title}</li>
            )}
          </ul>
        } */}

        <ul>
          {siteScriptResults && siteScriptResults.map(siteScript =>
            <li>{siteScript.Title}
              <button onClick={() => this._handleSiteScriptEdit(siteScript.Id)}>Edit</button>
              <button onClick={() => this._handleDeleteSiteScript(siteScript.Id)}>Delete</button>
            </li>
          )}
        </ul>

      </div>
    );
  }
}
