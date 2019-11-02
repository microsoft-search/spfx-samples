import * as React from 'react';
import styles from './GraphSearchApi.module.scss';
import { IGraphSearchApiProps } from './IGraphSearchApiProps';
import { IGraphSearchApiState } from './IGraphSearchApiState';
import { ClientMode } from './ClientMode';
import { ISearchResult } from './ISearchResult';
import * as strings from 'GraphSearchApiWebPartStrings';

import {
  autobind,
  PrimaryButton,
  TextField,
  Label,
  DetailsList,
  IColumn,
  IChoiceGroupOption,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  Checkbox,
  ChoiceGroupOption,
  Dropdown,
  ChoiceGroup
} from 'office-ui-fabric-react';

let _searchListColumns = [
  {
    key: 'link',
    name: 'link',
    fieldName: 'link',
    minWidth: 35,
    maxWidth: 35,
    isResizable: true
  },
  {
    key: 'id',
    name: 'ID',
    fieldName: 'id',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'type',
    name: 'Type',
    fieldName: 'type',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'summary',
    name: 'summary',
    fieldName: 'summary',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'source',
    name: 'source',
    fieldName: 'source',
    minWidth: 150,
    maxWidth: 300,
    isResizable: true
  },
  {
    key: 'score',
    name: 'score',
    fieldName: 'score',
    minWidth: 30,
    maxWidth: 30,
    isResizable: true
  },
  {
    key: 'sortField',
    name: 'sortField',
    fieldName: 'sortField',
    minWidth: 50,
    maxWidth: 50,
    isResizable: true
  }
];

import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";

function _renderItemColumn(item: ISearchResult, index: number, column: IColumn) {
  const fieldContent = item[column.fieldName as keyof ISearchResult] as string;

  switch (column.key) {
    case 'link':
      return (
        <a href={item.link} target="_blank">View</a>
        /*
        <div>
          {item.source.body.content}
        </div>
        */
      );
    case 'source':
      return (
        <div className={styles.hoverArea}>
        Result
        <div id={item.id} className={styles.hoverContent}>
          {item.source}
        </div>
        </div>
        
      );

      default:
      return <span>{fieldContent}</span>;
  }
}

export default class GraphSearchApi extends React.Component<IGraphSearchApiProps, IGraphSearchApiState> {

  constructor(props: IGraphSearchApiProps, state: IGraphSearchApiState) {
    
    super(props);

    // Initialize the state of the component
    this.state = {
      results: [],
      searchFor: "",
      externalPath: "",
      resultType: "message",
      includeEvents: false,
      includeFiles: false,
      includeMessages: false
    };
  }

  public render(): React.ReactElement<IGraphSearchApiProps> {

    const { clientMode } = this.props;
    
    return (

      <div className={ styles.graphSearchApi }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Search the Microsoft Graph</span>
              <p className={ styles.form }>
                <TextField 
                    label={ strings.SearchFor } 
                    required={ true } 
                    value={ this.state.searchFor }
                    onChanged={ this._onSearchForChanged }
                    onGetErrorMessage={ this._getSearchForErrorMessage }
                  />
                  <ChoiceGroup
                    defaultSelectedKey="message"
                    onChange={ this._onChoiceChanged }
                    label='Pick one'
                    required={ true }
                    options={ [ 
                      { 
                        key: 'message',
                        text: 'message', 
                      }, 
                      { 
                        key: 'driveItem',  
                        text: 'driveItem', 
                      }, 
                      { 
                        key: 'event',  
                        text: 'event', 
                      }, 
                      { 
                        key: 'externalItem',  
                        text: 'externalItem', 
                      }, 
                      { 
                        key: 'externalFile',  
                        text: 'externalFile', 
                      } 
                    ] }
                  />
                  <TextField 
                    label={ strings.ExternalPath } 
                    required={ false } 
                    value={ this.state.externalPath }
                    onChanged={ this._onExtPathChanged }
                    onGetErrorMessage={ this._getSearchForErrorMessage }
                  />
              </p>
              {
                (clientMode === ClientMode.aad || clientMode === ClientMode.graph) ?
                  <p className={styles.form}>
                    <PrimaryButton
                      text='Search'
                      title='Search'
                      onClick={this._search}
                    />
                  </p>
                  : <p>Configure client mode by editing web part properties.</p>
              }
              {
                (this.state.results != null && this.state.results.length > 0) ?
                  <p className={ styles.form }>
                  <DetailsList
                      items={ this.state.results }
                      columns={ _searchListColumns }
                      onRenderItemColumn={_renderItemColumn}
                      setKey='set'
                      checkboxVisibility={ CheckboxVisibility.hidden }
                      selectionMode={ SelectionMode.none }
                      layoutMode={ DetailsListLayoutMode.fixedColumns }
                      compact={ true }
                  />
                </p>
                : null
              }
            </div>
          </div>
        </div>
      </div>
    );
  }

  @autobind
  private _onChoiceChanged(ev: React.FormEvent<HTMLInputElement>, option: any): void { 
    this.setState({resultType: option.text}); 
  } 

  @autobind
  private _onExtPathChanged(newValue: string): void {

    // Update the component state accordingly to the current user's input
    this.setState({
      externalPath: newValue
    });
  }

  @autobind
  private _onSearchForChanged(newValue: string): void {

    // Update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue
    });
  }

  private _getSearchForErrorMessage(value: string): string {
    // The search for text cannot contain spaces
    return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }

  @autobind
  private _search(): void {

    console.log(this.props.clientMode);

    // Based on the clientMode value search users
    switch (this.props.clientMode)
    {
      case ClientMode.aad:
        this._searchWithAad();
        break;
      case ClientMode.graph:
        this._searchWithGraphSearch();
        break;
    }
  }

  private _searchWithAad(): void {

    // Log the current operation
    console.log("Using _searchWithAad() method");
  
    // Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        return client
          .get(
            `https://graph.microsoft.com/beta/search/query')`,
            AadHttpClient.configurations.v1
          );
      })
      .then(response => {
        return response.json();
      })
      .then(json => {
  
        var users: Array<ISearchResult> = new Array<ISearchResult>();

            // Map the JSON response to the output array
            json.value[0].hitsContainers[0].hits.map((item: any) => {
              users.push( { 
                id: item._id,
                type: item._source["@odata.type"],
                link: item._source.webLink,
                score: item._score,
                sortField: item._sortField,
                source: JSON.stringify(item._source),
                summary: item._summary
              });
            });
  
        // Update the component state accordingly to the result
        this.setState(
          {
            results: users,
          }
        );
      })
      .catch(error => {
        
        var users: Array<ISearchResult> = new Array<ISearchResult>();

        this.setState(
          {
            results: users,
          }
        );

        console.error(error);

      });
  }

  private _searchWithGraphSearch(): void {

    // Log the current operation
    console.log("Using _searchWithGraph() method");

    var query = this.state.searchFor;

    /*
    var types = [
      "microsoft.graph.message",
      "microsoft.graph.driveItem",
      "microsoft.graph.event"
    ];
    */

    var sources = [];

    switch (this.state.resultType) {
      case 'event':
        break;
        case 'message':
        break;
        case 'driveItem':
        break;
    }
    


    const data  = {
      requests: 
      [
        {
          entityTypes: [
            "microsoft.graph." + this.state.resultType,
          ],
          query: {
            query_string : {
              query : query
            }    
          },
          from: 0,
          size: 25,
          stored_fields: 
          [
            "from",
            "to",
            "subject",
            "body"
          ]
        }
    ]
  };
  
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api("search/query")
          .version("beta")
          .post(data)
          .then((res) => {

            console.log(res);        

            var users: Array<ISearchResult> = new Array<ISearchResult>();

            // Map the JSON response to the output array
            res.value[0].hitsContainers[0].hits.map((item: any) => {
              var link = "";
              if (item._source.webLink)
              {
                link = item._source.webLink;
              }
              if (item._source.webUrl)
              {
                link = item._source.webUrl;
              }
              if (this.state.resultType == 'event')
              {
                link = "https://outlook.office365.com/calendar/view/month";
              }
              users.push( { 
                id: item._id,
                type: item._source["@odata.type"],
                link: link,
                score: item._score,
                sortField: item._sortField,
                source: JSON.stringify(item._source),
                summary: item._summary
              });
            });

            // Update the component state accordingly to the result
            this.setState(
              {
                results: users,
              }
            );

          })
          .catch((err) => {

            var users: Array<ISearchResult> = new Array<ISearchResult>();
        
            this.setState(
              {
                results: users,
              }
            );

            console.error(err);

          });
      });
  }
}
