import * as React from 'react';
import styles from './SpSiteDesigner.module.scss';
import { ISpSiteDesignerProps } from '../types/ISpSiteDesignerProps';
import { ISpSiteDesignerState } from '../types/ISpSiteDesignerState';

import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';

import MonacoEditor from 'react-monaco-editor';
import DualListBox from 'react-dual-listbox';
import 'react-dual-listbox/lib/react-dual-listbox.css';

const config = {
  siteScripts: {
    perTenant: 30,
    characterLimit: 20000
  },
  siteDesigns: {
    perTenant: 20,
    actionLimit: 30
  }
};

export default class SpSiteDesigner extends React.Component<ISpSiteDesignerProps, ISpSiteDesignerState> {
  constructor(props: any) {
    super(props);
    this.state = {
      siteScriptResults: null,
      loading: true,
      siteScriptCharacterCount: 0,
      siteScriptForm: {
        title: "",
        content: "",
        description: ""
      },
      siteDesignForm: {
        title: "",
        description: "",
        webTemplate: "",
        previewImageUrl: "",
        previewImageAltText: "",
        selectedSiteScripts: []
      },
      siteScriptActionCount: 0
    };
    this._handleInputChange = this._handleInputChange.bind(this);
    this._handleEditorChange = this._handleEditorChange.bind(this);
    this._handleSiteScriptFormSubmit = this._handleSiteScriptFormSubmit.bind(this);
    this._handleSiteDesignFormSubmit = this._handleSiteDesignFormSubmit.bind(this);
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
      if (response.status === 204) {
        return {};
      } else {
        return response.json();
      }
    });
  }

  private _saveSiteScript(): any {
    const siteScriptData = JSON.parse(this.state.siteScriptForm.content);
    if (this.state.selectedSiteScriptID) {
      // Update Site Script
      return this._restRequest(
        `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`, {
          updateInfo: {
            Id: this.state.selectedSiteScriptID,
            Title: this.state.siteScriptForm.title,
            Content: this.state.siteScriptForm.content,
            Description: this.state.siteScriptForm.description
          }
        }
      ).then(() => {
        this._loadData();
      });
    }
    // Create Site Script
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title='${this.state.siteScriptForm.title}'`,
      siteScriptData
    ).then(() => {
      this._loadData();
    });
  }

  private _getSiteScripts(): Promise<any> {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`
    ).then((response) => response);
  }

  private _getSiteScriptMetadata(id: string): any {
    this._clearSiteScriptForm();
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata`,
      id
    ).then((response) => {
      const siteScriptCharacterCount: number = this._countSiteScriptCharacters(response.Content);
      this.setState({
        siteScriptCharacterCount: siteScriptCharacterCount,
        selectedSiteScriptID: response.Id,
        siteScriptForm: {
          title: response.Title,
          content: response.Content,
          description: response.Description
        }
      });
    });
  }

  private _deleteSiteScript(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript`,
      id
    ).then(() => {
      this._loadData();
    });
  }

  private _saveSiteDesign(): any {
    if (this.state.selectedSiteDesignID) {
      // Update Site Design
      return this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign", {
        updateInfo: {
          Title: this.state.siteDesignForm.title,
          Id: this.state.selectedSiteDesignID,
          Description: this.state.siteDesignForm.description,
          SiteScriptIds: this.state.siteDesignForm.selectedSiteScripts,
          WebTemplate: this.state.siteDesignForm.webTemplate,
          PreviewImageUrl: this.state.siteDesignForm.previewImageUrl,
          PreviewImageAltText: this.state.siteDesignForm.previewImageAltText
        }
      }).then((response) => {
        this._loadData();
      });
    }
    // Create Site Design
    return this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign", {
      info: {
        Title: this.state.siteDesignForm.title,
        Description: this.state.siteDesignForm.description,
        SiteScriptIds: this.state.siteDesignForm.selectedSiteScripts,
        WebTemplate: this.state.siteDesignForm.webTemplate,
        PreviewImageUrl: this.state.siteDesignForm.previewImageUrl,
        PreviewImageAltText: this.state.siteDesignForm.previewImageAltText
      }
    }).then((response) => {
      this._loadData();
    });
  }

  private _deleteSiteDesign(id: string): any {
    return this._restRequest(
      `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign`,
      id
    ).then((response) => {
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

      this._countSiteScriptActions(response.SiteScriptIds);

      this.setState({
        selectedSiteDesignID: response.Id,
        siteDesignForm: {
          title: response.Title,
          description: response.Description,
          selectedSiteScripts: response.SiteScriptIds,
          webTemplate: response.WebTemplate,
          previewImageUrl: response.PreviewImageUrl,
          previewImageAltText: response.PreviewImageAltText
        }
      });
    });
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

  public _handleInputChange = (form, name) => value => {
    this.setState(state => ({
      [form]: {
        ...state[form],
        [name]: value
      }
    }));
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

  public _countSiteScriptCharacters(siteScript) {
    siteScript = siteScript.split("");
    return siteScript.length;
  }

  public _handleEditorChange(newValue, e) {
    const siteScriptCharacterCount: number = this._countSiteScriptCharacters(newValue);
    this.setState(state => ({
      siteScriptCharacterCount: siteScriptCharacterCount,
      siteScriptForm: {
        ...state.siteScriptForm,
        content: newValue
      }
    }));
  }

  public _handleSiteScriptFormSubmit(event) {
    event.preventDefault();
    this._saveSiteScript();
  }

  public _handleSiteDesignFormSubmit(event) {
    event.preventDefault();
    this._saveSiteDesign();
  }

  public _handleNewSiteScriptClick() {
    this.setState({
      selectedSiteScriptID: null
    });
    this._clearSiteScriptForm();
  }

  public _handleNewSiteDesignClick() {
    this.setState({
      selectedSiteDesignID: null
    });
    this._clearSiteDesignForm();
  }

  public _clearSiteScriptForm() {
    this.setState({
      siteScriptForm: {
        title: "",
        content: "",
        description: ""
      }
    });
  }

  public _clearSiteDesignForm() {
    this.setState({
      siteScriptActionCount: 0,
      siteDesignForm: {
        title: "",
        description: "",
        webTemplate: "",
        previewImageUrl: "",
        previewImageAltText: "",
        selectedSiteScripts: []
      }
    });
  }

  public _countSiteScriptActions(siteScripts) {
    let requestList = [];

    for (let i: number = 0; i < siteScripts.length; i++) {
      const siteScript: any = {
        id: siteScripts[i]
      };
      requestList.push(this._restRequest(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata`, siteScript));
    }

    Promise.all(requestList)
      .then((response) => {
        let actionCount: number = 0;
        let subactionsCount: number = 0;
        for (let i: number = 0; i < response.length; i++) {
          const siteScript = response[i];
          const siteScriptObj = JSON.parse(siteScript.Content);
          for (let j: number = 0; j < siteScriptObj.actions.length; j++) {
            let action = siteScriptObj.actions[j];
            if ('subactions' in action) {
              subactionsCount += action.subactions.length;
            }
          }
          actionCount += siteScriptObj.actions.length;
          actionCount += subactionsCount;
        }
        this.setState({
          siteScriptActionCount: actionCount
        });
      });
  }

  public _actionCountFormat(siteScriptActionCount: number): string {
    let actionCountFormat;
    if (siteScriptActionCount) {
      if (siteScriptActionCount <= 25) {
        actionCountFormat = styles.green;
      }
      if (siteScriptActionCount >= 26 && siteScriptActionCount <= 28) {
        actionCountFormat = styles.yellow;
      }
      if (siteScriptActionCount >= 29) {
        actionCountFormat = styles.red;
      }
    }
    return actionCountFormat;
  }

  public render(): React.ReactElement<ISpSiteDesignerProps> {

    const { loading, siteScriptResults, siteDesignResults, siteDesignForm, siteScriptForm, selectedSiteScriptID, selectedSiteDesignID, siteScriptActionCount, siteScriptCharacterCount } = this.state;

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

    return (
      <div className={styles.spSiteDesigner}>
        <Pivot>
          <PivotItem headerText="Site Scripts">
            <div className={styles.container}>
              <div className={styles.row}>
                <DefaultButton text="Create a New Site Script" primary={true} onClick={() => this._handleNewSiteScriptClick()} />
              </div>
              <div className={styles.row}>
                <div className={styles.sidebar}>
                  <div>
                    <div className={`${styles.dFlex} ${styles.alignItemsCenter} ${styles.justifyContentBetween}`}>
                      <h2 className={styles.sidebarTitle}>Your Site Scripts</h2>
                      <span className={styles.sidebarLimitCount}>
                        {siteDesignResults && siteScriptResults.length} / {config.siteScripts.perTenant}
                      </span>
                    </div>
                    <ul className={styles.sidebarList}>
                      {siteScriptResults && siteScriptResults.map(siteScript =>
                        <li key={siteScript.Id} title={siteScript.Title} onClick={() => this._handleSiteScriptEdit(siteScript.Id)} className={(this.state.selectedSiteScriptID === siteScript.Id && styles.selected)}>
                          <div className={styles.listLabel}>
                            {siteScript.Title}
                          </div>
                        </li>
                      )}
                    </ul>
                  </div>
                </div>
                <div className={`${styles.main} ${styles.pl3}`}>
                  <div>
                    <h2 className={styles.formTitle}>{(selectedSiteScriptID ? "Edit" : "Create")} Site Script</h2>
                    <form onSubmit={this._handleSiteScriptFormSubmit}>
                      <TextField label="Title" value={siteScriptForm.title} onChanged={this._handleInputChange('siteScriptForm', 'title')} />
                      {selectedSiteScriptID && <TextField label="Description" value={siteScriptForm.description} onChanged={this._handleInputChange('siteScriptForm', 'description')} />}
                      <div>


                        <div className={`${styles.dFlex} ${styles.alignItemsCenter} ${styles.justifyContentBetween}`}>
                          <div className={styles.p5}>JSON</div>
                          <span className={styles.sidebarLimitCount}>
                            {siteScriptResults && siteScriptCharacterCount} / {config.siteScripts.characterLimit}
                          </span>
                        </div>


                        <MonacoEditor
                          width="100%"
                          height="300"
                          language="json"
                          theme="vs-dark"
                          value={siteScriptForm.content}
                          options={options}
                          onChange={this._handleEditorChange}
                          editorDidMount={this.editorDidMount}
                        />
                      </div>
                      <div className={`${styles.dFlex} ${styles.justifyContentEnd}`}>
                        {selectedSiteScriptID && <DefaultButton text="Delete" onClick={() => this._handleDeleteSiteScript(this.state.selectedSiteScriptID)} className={styles.mt3} />}
                        <DefaultButton text="Save" type="Submit" primary={true} className={`${styles.mt3} ${styles.ml2}`} />
                      </div>
                    </form>
                  </div>
                </div>
              </div>
            </div>
          </PivotItem>
          <PivotItem headerText="Site Designs">
            <div className={styles.container}>
              <div className={styles.row}>
                <DefaultButton text="Create a New Site Design" primary={true} onClick={() => this._handleNewSiteDesignClick()} />
              </div>
              <div className={styles.row}>
                <div className={styles.sidebar}>
                  <div>
                    <div className={`${styles.dFlex} ${styles.alignItemsCenter} ${styles.justifyContentBetween}`}>
                      <h2 className={styles.sidebarTitle}>Your Site Designs</h2>
                      <span className={styles.sidebarLimitCount}>
                        {siteDesignResults && siteDesignResults.length} / {config.siteDesigns.perTenant}
                      </span>
                    </div>
                    <ul className={styles.sidebarList}>
                      {siteDesignResults && siteDesignResults.map(siteDesign =>
                        <li key={siteDesign.Id} title={siteDesign.Title} onClick={() => this._handleSiteDesignEdit(siteDesign.Id)} className={(this.state.selectedSiteDesignID === siteDesign.Id && styles.selected)}>
                          <div className={styles.listLabel}>
                            {siteDesign.Title}
                          </div>
                        </li>
                      )}
                    </ul>
                  </div>
                </div>
                <div className={`${styles.main} ${styles.pl3}`}>
                  <div>
                    <div className={`${styles.dFlex} ${styles.justifyContentBetween} ${styles.alignItemsCenter}`}>
                      <h2 className={styles.formTitle}>{(selectedSiteDesignID ? "Edit" : "Create")} Site Design</h2>
                      <div className={`${styles.actionCount} ${this._actionCountFormat(siteScriptActionCount)}`}>
                        <span className={styles.actionCountLabel}>Actions:</span>
                        <span className={styles.actionCountValue}>{siteScriptActionCount}</span>/<span className={styles.actionLimit}>{config.siteDesigns.actionLimit}</span>
                      </div>
                    </div>
                    <form onSubmit={this._handleSiteDesignFormSubmit}>
                      <TextField label="Title" value={siteDesignForm.title} onChanged={this._handleInputChange('siteDesignForm', 'title')} />
                      <TextField label="Description" value={siteDesignForm.description} onChanged={this._handleInputChange('siteDesignForm', 'description')} />
                      <TextField label="Web Template" value={siteDesignForm.webTemplate} onChanged={this._handleInputChange('siteDesignForm', 'webTemplate')} />
                      <TextField label="Preview Image URL" value={siteDesignForm.previewImageUrl} onChanged={this._handleInputChange('siteDesignForm', 'previewImageUrl')} />
                      <TextField label="Preview Image Alt Text" value={siteDesignForm.previewImageAltText} onChanged={this._handleInputChange('siteDesignForm', 'previewImageAltText')} />
                      <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                        <div><h4>Available Site Scripts</h4></div>
                        <div><h4>Added to Site Design</h4></div>
                      </div>
                      {siteScriptOptions && <DualListBox
                        availableLabel="Available Site Scripts"
                        selectedLabel="Added to Site Design"
                        simpleValue={true}
                        options={siteScriptOptions}
                        selected={siteDesignForm.selectedSiteScripts}
                        onChange={(selectedSiteScripts) => {
                          this.setState(state => ({
                            siteDesignForm: {
                              ...state.siteDesignForm,
                              selectedSiteScripts
                            }
                          }));
                        }}
                      />}
                      <div className={`${styles.dFlex} ${styles.justifyContentEnd}`}>
                        {selectedSiteDesignID && <DefaultButton text="Delete" onClick={() => this._handleDeleteSiteDesign(this.state.selectedSiteDesignID)} className={styles.mt3} />}
                        <DefaultButton text="Save" type="Submit" primary={true} className={`${styles.mt3} ${styles.ml2}`} />
                      </div>
                    </form>
                  </div>
                </div>
              </div>
            </div>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}