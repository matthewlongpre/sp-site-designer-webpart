import * as React from 'react';
import styles from './SpSiteDesigner.module.scss';
import { ISpSiteDesignerProps } from './ISpSiteDesignerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';

import MonacoEditor from 'react-monaco-editor';

import DualListBox from 'react-dual-listbox';
import 'react-dual-listbox/lib/react-dual-listbox.css';

import SiteScriptForm from './SiteScriptForm';
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
  selectedSiteScriptID?: any;
  editingSelectedSiteScriptTitle?: string;
  editingSelectedSiteScriptContent?: string;
  loading?: boolean;
}

export default class SpSiteDesigner extends React.Component<ISpSiteDesignerProps, ISpSiteDesignerState> {
  constructor(props: any) {
    super(props);
    this.state = {
      siteScriptData: null,
      siteScriptResults: null,
      siteScriptTitle: null,
      selectedSiteScripts: [],
      loading: true,
      editingSelectedSiteScriptTitle: "",
      editingSelectedSiteScriptContent: ""
    };
    this._handleInputChange = this._handleInputChange.bind(this);
    this._handleCreateSiteScriptClick = this._handleCreateSiteScriptClick.bind(this);
    this._handleEditorChange = this._handleEditorChange.bind(this);
  }

  public componentDidMount() {
    this._loadData();
  }

  public _loadData() {
    this.setState({
      loading: true
    });
    Promise.all([this._getSiteScripts(), this._getSiteDesigns()])
      .then((response) => {
        const [siteScriptResults, siteDesignResults] = response;
        this.setState({
          loading: false,
          siteScriptResults: siteScriptResults.value,
          siteDesignResults: siteDesignResults.value
        });
      });
  }

  public _reset() {
    this.setState({
      selectedSiteDesignID: null,
      selectedSiteScriptID: null
    });
  }

  public baseUrl: string = '/';

  private _getUrl(relativeUrl: string): string {
    return (this.baseUrl + '//' + relativeUrl).replace(/\/{2,}/, '/');
  }

  private _restRequest(url: string, params: any = null): Promise<any> {
    const restUrl = this._getUrl(url);
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

  private _saveSiteScript(siteScriptTitle: string, siteScriptData: string): any {
    siteScriptData = JSON.parse(siteScriptData);
    if (this.state.selectedSiteScriptID) {
      // Update Site Script
      return this._restRequest(
        `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`, {
          updateInfo: {
            Id: this.state.selectedSiteScriptID,
            Title: this.state.editingSelectedSiteScriptTitle,
            Content: this.state.editingSelectedSiteScriptContent
          }
        }
      );      
    }
    // Create Site Script
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title='${siteScriptTitle}'`,
      siteScriptData
    );
  }

  private _getSiteScripts(): Promise<any> {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`
    ).then((response) => response);
  }

  private _getSiteScriptMetadata(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata`,
      id
    ).then((response) => {
      this.setState({
        selectedSiteScriptID: response.Id,
        editingSelectedSiteScriptTitle: response.Title,
        editingSelectedSiteScriptContent: response.Content,
      });
    });
  }

  private _deleteSiteScript(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript`,
      id
    ).then((response) => response);
  }

  private _saveSiteDesign(): any {
    if (this.state.selectedSiteDesignID) {
      // Update Site Design
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
    // Create Site Design
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
    ).then((response) => response);
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

  private _handleCreateSiteScriptClick(): any {
    this._saveSiteScript(this.state.editingSelectedSiteScriptTitle, this.state.editingSelectedSiteScriptContent);
  }

  private _handleCreateSiteDesignClick(): any {
    this._saveSiteDesign();
  }

  private _handleGetSiteScriptClick(): any {
    this._getSiteScripts();
  }

  private _handleGetSiteDesignClick(): any {
    this._getSiteDesigns();
  }

  private _handleResetClick(): any {
    this._reset();
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
    };
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

  public editorDidMount(editor, monaco) {
    editor.focus();
  }
  public _handleEditorChange(newValue, e) {
    this.setState({
      editingSelectedSiteScriptContent: newValue
    });
  }

  public render(): React.ReactElement<ISpSiteDesignerProps> {

    const { loading, siteScriptResults, siteDesignResults, selectedSiteScripts } = this.state;

    const options = {
      selectOnLineNumbers: true
    };

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

    if (loading) return <div></div>;

    return (
      <div className={styles.spSiteDesigner} >

        {/* <button onClick={() => this._handleGetSiteScriptClick()}>Get Site Scripts</button>
        <button onClick={() => this._handleGetSiteDesignClick()}>Get Site Designs</button>
        <button onClick={() => this._handleResetClick()}>Reset</button> */}

        <div>
          <h2>Site Script</h2>
          <form>
            <div><div>Title</div><input id="siteScriptTitle" name="editingSelectedSiteScriptTitle" value={this.state.editingSelectedSiteScriptTitle} onChange={this._handleInputChange}></input></div>
            <div><div>JSON</div>

              <MonacoEditor
                width="100%"
                height="300"
                language="json"
                theme="vs-dark"
                value={this.state.editingSelectedSiteScriptContent}
                options={options}
                onChange={this._handleEditorChange}
                editorDidMount={this.editorDidMount}
              />

              {/* <textarea style={{width:'400px'}} id="siteScriptData" name="editingSelectedSiteScriptContent" value={this.state.editingSelectedSiteScriptContent} onChange={this._handleInputChange}></textarea> */}

            </div>
          </form>
          <button onClick={this._handleCreateSiteScriptClick}>Save Site Script</button>
        </div>

        {/* <SiteScriptForm handleCreateSiteScriptClick={this._handleCreateSiteScriptClick} initialState={this.state.editingSelectedSiteScript} /> */}

        <div>
          <h2>Site Design</h2>
          <form>
            <div><div>Title</div><input name="siteDesignTitle" value={this.state.siteDesignTitle} onChange={this._handleInputChange} /></div>
            <div><div>Description</div><input name="siteDesignDescription" value={this.state.siteDesignDescription} onChange={this._handleInputChange} /></div>
            <div><div>Web Template</div><input name="siteDesignWebTemplate" value={this.state.siteDesignWebTemplate} onChange={this._handleInputChange} /></div>
            <div><div>Preview Image URL</div><input name="siteDesignPreviewImageUrl" value={this.state.siteDesignPreviewImageUrl} onChange={this._handleInputChange} /></div>
            <div><div>Preview Image Alt Text</div><input name="siteDesignPreviewImageAltText" value={this.state.siteDesignPreviewImageAltText} onChange={this._handleInputChange} /></div>

            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <div><h4>Available Site Scripts</h4></div>
              <div><h4>Added to Site Design</h4></div>
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
          <button onClick={() => this._handleCreateSiteDesignClick()}>Save Site Design</button>
        </div>

        <div>
          <h2>Available Site Designs</h2>
          <ul>
            {siteDesignResults && siteDesignResults.map(siteDesign =>
              <li>
                <div>
                  {siteDesign.Title}
                </div>
                <div>
                  <button onClick={() => this._handleSiteDesignEdit(siteDesign.Id)}>Edit</button>
                  <button onClick={() => this._handleDeleteSiteDesign(siteDesign.Id)}>Delete</button>
                </div>
              </li>
            )}
          </ul>
        </div>

        <div>
          <h2>Available Site Scripts</h2>
          <ul>
            {siteScriptResults && siteScriptResults.map(siteScript =>
              <li>
                <div>
                  {siteScript.Title}
                </div>
                <div>
                  <button onClick={() => this._handleSiteScriptEdit(siteScript.Id)}>Edit</button>
                  <button onClick={() => this._handleDeleteSiteScript(siteScript.Id)}>Delete</button>
                </div>
              </li>
            )}
          </ul>
        </div>
      </div>
    );
  }
}