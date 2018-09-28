import * as React from 'react';

export interface ISiteScriptFormProps {
  handleCreateSiteScriptClick: any;
  initialState: any;
}

export interface ISiteScriptFormState {
  title?: any;
  content?: any;
}

export default class SiteScriptForm extends React.Component<ISiteScriptFormProps, ISiteScriptFormState> {
  constructor(props: any) {
    super(props);
    this.state = {
      title: this.props.initialState.Title,
      content: this.props.initialState.Content
    };
    this._handleInputChange = this._handleInputChange.bind(this);
    this._handleCreateSiteScriptClick = this._handleCreateSiteScriptClick.bind(this);
  }
  private _handleInputChange(event: any): any {
    console.log('change');
    const target = event.target;
    const value = target.type === 'checkbox' ? target.checked : target.value;
    const name = target.name;
    this.setState({
      [name]: value
    });
  }
  public _handleCreateSiteScriptClick() {
    const siteScript = {
      siteScriptTitle: this.state.title,
      siteScriptData: this.state.content
    };
    this.props.handleCreateSiteScriptClick(siteScript);
  }
  public _componentDidUpdate() {
    console.log('updated');

  }
  public render() {
    console.log(this.props.initialState);
    return (
      <div>
        <h2>Site Script</h2>
        <form>
          <div><div>Title</div><input id="siteScriptTitle" name="title" value={this.state.title} onChange={this._handleInputChange}></input></div>
          <div><div>JSON</div><textarea id="siteScriptData" name="content" value={this.state.content} onChange={this._handleInputChange}></textarea></div>
        </form>
        <button onClick={this._handleCreateSiteScriptClick}>Save Site Script</button>
      </div>
    );
  } 
}